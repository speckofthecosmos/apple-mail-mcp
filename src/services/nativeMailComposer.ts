import { spawnSync } from "node:child_process";
import {
  MAX_NATIVE_CLIPBOARD_SNAPSHOT_BYTES,
  MAX_NATIVE_EDITOR_OUTPUT_BYTES,
  recoverNativeClipboard,
} from "./nativeClipboardRecovery.js";
import { randomUUID } from "node:crypto";
import { basename, isAbsolute } from "node:path";

export interface NativeEditorResult {
  success: boolean;
  error?: string;
}

export interface NativeAuthoredBody {
  messageId: number;
  markerSubject: string;
  body: string;
  expectedSender: string;
  /** Absolute, policy-checked paths of files to paste below the authored prose. */
  attachments?: string[];
}

/** Default bound on one native-editor run; see nativeEditorTimeoutMs. */
export const NATIVE_EDITOR_DEFAULT_TIMEOUT_MS = 60000;
/** The former fixed bound, kept as the floor for a configured value. */
export const NATIVE_EDITOR_MIN_TIMEOUT_MS = 15000;
/** Extra time per pasted attachment: archiving it and waiting for it to appear. */
export const NATIVE_EDITOR_ATTACHMENT_TIMEOUT_MS = 5000;

/**
 * Wall-clock bound for one native-editor subprocess. Accessibility reads cost
 * roughly 20 ms each in Mail's composer, so a reply with a long quoted history
 * and HTML signature can spend tens of seconds just locating and reading back
 * the editor (28.7 s measured live). `APPLE_MAIL_MCP_NATIVE_EDITOR_TIMEOUT_MS`
 * overrides the base (minimum 15000); attachments add a fixed allowance each.
 */
export function nativeEditorTimeoutMs(attachments = 0): number {
  const configured = Number(process.env.APPLE_MAIL_MCP_NATIVE_EDITOR_TIMEOUT_MS);
  const base =
    Number.isFinite(configured) && configured > 0
      ? Math.max(NATIVE_EDITOR_MIN_TIMEOUT_MS, Math.floor(configured))
      : NATIVE_EDITOR_DEFAULT_TIMEOUT_MS;
  return base + attachments * NATIVE_EDITOR_ATTACHMENT_TIMEOUT_MS;
}

/** New prose only: Mail owns the signature, quotation, and reply headers. */
export function renderNativeAuthoredBody(body: string): string {
  const entities: Record<string, string> = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  };
  return body
    .normalize("NFC")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map(
      (line) =>
        `<div style="font-size: inherit;">${line.replace(/[&<>"']/g, (c) => entities[c]) || "<br>"}</div>`
    )
    .join("");
}

// A fixed program, never caller-interpolated code. Payloads travel on stdin,
// not argv, shell commands, temporary files, or debug logs. No send command is
// present: the manager alone owns save/send after this bridge succeeds.
const NATIVE_EDITOR_SCRIPT = String.raw`
ObjC.import('AppKit');
ObjC.import('ApplicationServices');
ObjC.import('CoreServices');
ObjC.bindFunction('AEDeterminePermissionToAutomateTarget', ['int', ['void *', 'unsigned int', 'unsigned int', 'bool']]);
var currentStage = 'startup', focusFailure = null, targetFailure = null;
function stage(value) { currentStage = value; console.log('NATIVE_STAGE:' + value); }
function emit(record) {
  var data = $(JSON.stringify(record) + '\n').dataUsingEncoding($.NSUTF8StringEncoding);
  $.NSFileHandle.fileHandleWithStandardOutput.writeData(data);
}
function fail(code) { throw new Error(code); }
function preflight() {
  if (!$.AXIsProcessTrusted()) fail('accessibility');
  ['com.apple.mail', 'com.apple.systemevents'].forEach(function (id) {
    if (id === 'com.apple.systemevents' &&
        Number($.NSRunningApplication.runningApplicationsWithBundleIdentifier(id).count) === 0) {
      // LaunchServices launch is not an Apple Event and does not request TCC
      // permission. A stopped System Events otherwise reports procNotFound,
      // which must not be misreported as denied Automation permission.
      try {
        var workspace = $.NSWorkspace.sharedWorkspace;
        workspace.launchApplicationAtURLOptionsConfigurationError(
          workspace.URLForApplicationWithBundleIdentifier(id),
          $.NSWorkspaceLaunchWithoutActivation, $.NSDictionary.dictionary, Ref());
      } catch (_) { fail('automation-unavailable'); }
    }
    var target = $.NSAppleEventDescriptor.descriptorWithBundleIdentifier(id);
    // core/getd, askUserIfNeeded=false: never display an authorization prompt.
    var status;
    for (var attempt = 0; attempt < 12; attempt++) {
      status = $.AEDeterminePermissionToAutomateTarget(target.aeDesc, 0x636f7265, 0x67657464, false);
      if (status !== -600 || id !== 'com.apple.systemevents') break;
      if (attempt < 11) delay(0.15);
    }
    if (status === -1743 || status === -1744) fail('automation');
    if (status !== 0) fail('automation-unavailable');
  });
}
function attr(element, name) {
  try { return element.attributes.byName(name).value(); } catch (_) { return null; }
}
// Depth-first pre-order, which is document order: bodyText's readback compares
// the editor's text by position (new prose first, Mail's content last). A
// breadth-first walk orders by depth instead, so a shallow quoted line (Mail's
// "<file.txt>" placeholder for a quoted attachment) was read before the prose.
function descendants(root) {
  var result = [], pending = [root];
  while (pending.length) {
    var element = pending.pop();
    result.push(element);
    if (result.length > 2000) fail('target');
    var children = [];
    try { children = element.uiElements(); } catch (_) {}
    for (var i = children.length - 1; i >= 0; i--) pending.push(children[i]);
  }
  return result;
}
function findMarkedComposer(process, input, checkMessage) {
  // Mail can publish the marked outgoing message before its accessibility
  // window, subject, and editor have all caught up. Retry only observations;
  // never create, click, paste, or select an unmarked or ambiguous composer.
  for (var attempt = 0; attempt < 20; attempt++) {
    checkMessage();
    stage('target-window');
    var windows = process.windows().filter(function (window) {
      // Never walk the main mailbox tree: it may contain thousands of rows.
      return window.name() === input.markerSubject;
    });
    if (windows.length !== 1) {
      targetFailure = windows.length === 0 ? 'window-missing' : 'window-ambiguous';
    } else {
      var window = windows[0], elements = descendants(window);
      var subjects = elements.filter(function (element) {
        return attr(element, 'AXIdentifier') === 'Mail.subjectField' &&
          attr(element, 'AXValue') === input.markerSubject;
      });
      if (subjects.length !== 1) {
        targetFailure = subjects.length === 0 ? 'subject-missing' : 'subject-ambiguous';
      } else {
        stage('target-body');
        var bodies = elements.filter(function (element) {
          return attr(element, 'AXRole') === 'AXWebArea' &&
            attr(element, 'AXDescription') === 'message body';
        });
        if (bodies.length === 1) {
          targetFailure = null;
          return {window: window, subject: subjects[0], body: bodies[0]};
        }
        targetFailure = bodies.length === 0 ? 'body-missing' : 'body-ambiguous';
      }
    }
    if (attempt < 19) delay(0.1);
  }
  fail('target');
}
function normalized(value) { return String(value).normalize('NFC').replace(/\s+/g, ' ').trim(); }
function verifiedInsertion(before, after, authored, attachments) {
  // Attachments surface in the editor's text as object-replacement characters.
  function clean(value) { return normalized(String(value).replace(/\uFFFC/g, ' ')); }
  var oldText = clean(before), newText = clean(after), added = clean(authored);
  if (!attachments) {
    return added.length > 0 && newText.length > oldText.length &&
      newText === normalized(added + '\n' + oldText);
  }
  // Pasted files sit between the prose and Mail's existing signature/history and
  // may also expose a file label, so require the prose first and every
  // pre-existing character last, unchanged. Arrival is counted separately.
  return newText.length >= oldText.length + added.length &&
    newText.slice(0, added.length) === added &&
    newText.slice(newText.length - oldText.length) === oldText;
}
function bodyText(body) {
  var direct = attr(body, 'AXValue');
  if (typeof direct === 'string' && direct.length) return direct;
  return descendants(body).filter(function (element) {
    var role = attr(element, 'AXRole');
    return role === 'AXStaticText' || role === 'AXTextArea';
  }).map(function (element) {
    var value = attr(element, 'AXValue');
    if (typeof value !== 'string') fail('readback');
    return value;
  }).join('\n');
}
function address(value) {
  var text = String(value).trim(), match = text.match(/<([^<>]+)>$/);
  return (match ? match[1] : text).toLowerCase();
}
function escapeAttribute(value) {
  return String(value).replace(/[&<>"']/g, function (c) {
    return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c];
  });
}
function mimeType(name) {
  var dot = name.lastIndexOf('.');
  if (dot > 0) {
    try {
      // UniformTypeIdentifiers has no JXA bridge; message the class directly.
      var type = $.NSClassFromString('UTType').typeWithFilenameExtension(name.slice(dot + 1));
      var mime = type.isNil() ? null : ObjC.unwrap(type.preferredMIMEType);
      if (typeof mime === 'string' && mime) return mime;
    } catch (_) {}
  }
  return 'application/octet-stream';
}
function dictionary(pairs) {
  var result = $.NSMutableDictionary.alloc.init;
  Object.keys(pairs).forEach(function (key) { result.setObjectForKey(pairs[key], key); });
  return result;
}
// One com.apple.webarchive carries the prose and every file in a single paste.
// Mail inserts each <attachment> from its subresource, whose data must be a
// keyed-archived NSFileWrapper: raw file bytes display but fail at save. (A
// separate file-URL item in the same paste makes Mail drop the text instead.)
function webArchive(html, files) {
  var markup = html, resources = $.NSMutableArray.alloc.init;
  files.forEach(function (file) {
    if (!file || typeof file.path !== 'string' || file.path.charAt(0) !== '/' ||
        typeof file.name !== 'string' || !file.name) fail('attachments');
    var data = $.NSData.dataWithContentsOfFile(file.path);
    if (data.isNil()) fail('attachments');
    var wrapper = $.NSFileWrapper.alloc.initRegularFileWithContents(data);
    wrapper.preferredFilename = file.name;
    var archived = $.NSKeyedArchiver.archivedDataWithRootObjectRequiringSecureCodingError(wrapper, false, null);
    if (archived.isNil()) fail('attachments');
    var id = ObjC.unwrap($.NSUUID.UUID.UUIDString), mime = mimeType(file.name);
    markup += '<div><attachment webkitattachmentid="' + id + '" title="' + escapeAttribute(file.name) +
      '" type="' + escapeAttribute(mime) + '"></attachment></div>';
    resources.addObject(dictionary({
      WebResourceURL: $('applewebdata://attachment/' + id),
      WebResourceMIMEType: $(mime),
      WebResourceData: archived
    }));
  });
  var archive = $.NSPropertyListSerialization.dataWithPropertyListFormatOptionsError(dictionary({
    WebMainResource: dictionary({
      WebResourceData: $(markup).dataUsingEncoding($.NSUTF8StringEncoding),
      WebResourceMIMEType: $('text/html'),
      WebResourceTextEncodingName: $('UTF-8'),
      WebResourceURL: $('about:blank'),
      WebResourceFrameName: $('')
    }),
    WebSubresources: resources
  }), $.NSPropertyListBinaryFormat_v1_0, 0, null);
  if (archive.isNil()) fail('attachments');
  return archive;
}
// Mail's scripting attachment count reads a quote-less copy of the message that
// never sees pasted files. The editor exposes each attachment as an AXButton
// described "<file name>, <size>", e.g. "report.pdf, 2.4 MB".
function attachmentButtons(body, names) {
  var counts = {};
  names.forEach(function (name) { counts[name] = 0; });
  descendants(body).forEach(function (element) {
    if (attr(element, 'AXRole') !== 'AXButton') return;
    var description = attr(element, 'AXDescription');
    if (typeof description !== 'string') return;
    description = description.normalize('NFC');
    Object.keys(counts).forEach(function (name) {
      if (description.indexOf(name.normalize('NFC') + ', ') === 0) counts[name]++;
    });
  });
  return counts;
}
function waitForAttachments(body, names, baseline) {
  var expected = {};
  names.forEach(function (name) {
    expected[name] = (expected.hasOwnProperty(name) ? expected[name] : baseline[name]) + 1;
  });
  var deadline = Date.now() + 10000;
  for (;;) {
    var counts = attachmentButtons(body, names);
    if (Object.keys(expected).every(function (name) { return counts[name] >= expected[name]; })) return;
    if (Date.now() >= deadline) fail('attachment-arrival');
    delay(0.25);
  }
}
function writeItems(clipboard, items) {
  var array = $.NSMutableArray.alloc.init;
  items.forEach(function (item) { array.addObject(item); });
  return clipboard.writeObjects(array);
}
function snapshotClipboard(clipboard, token) {
  var limit = ${MAX_NATIVE_CLIPBOARD_SNAPSHOT_BYTES};
  function bytes(value) { return Number($(JSON.stringify(value)).lengthOfBytesUsingEncoding($.NSUTF8StringEncoding)); }
  var originalCount = Number(clipboard.changeCount);
  var recovery = {kind: 'clipboard-recovery', items: [], originalCount: originalCount, token: token};
  var size = bytes(recovery), items = clipboard.pasteboardItems;
  for (var i = 0; i < Number(items.count); i++) {
    var item = items.objectAtIndex(i), types = item.types, recoveryTypes = [];
    var saved = $.NSPasteboardItem.alloc.init;
    size += 2 + (i ? 1 : 0); // Item brackets and comma, inside the existing [].
    if (size > limit) fail('clipboard');
    for (var j = 0; j < Number(types.count); j++) {
      var type = types.objectAtIndex(j), data = item.dataForType(type);
      // Preserve every flavor, including binary data; never silently drop a
      // lazy/unrepresentable flavor. Account for base64 before allocating it.
      if (!data) fail('clipboard');
      var entry = {type: ObjC.unwrap(type), data: ''};
      size += bytes(entry) + 4 * Math.ceil(Number(data.length) / 3) + (j ? 1 : 0);
      if (size > limit) fail('clipboard');
      if (!saved.setDataForType(data, type)) fail('clipboard');
      entry.data = ObjC.unwrap(data.base64EncodedStringWithOptions(0));
      recoveryTypes.push(entry);
    }
    recovery.items.push(recoveryTypes);
  }
  if (Number(clipboard.changeCount) !== originalCount || bytes(recovery) > limit) fail('clipboard');
  return recovery;
}
function main(input) {
  stage('preflight');
  preflight();
  if (input.operation === 'preflight') return;
  var mail = Application('com.apple.mail');
  var events = Application('com.apple.systemevents');
  var process = events.processes.byName('Mail');
  function checkMessage() {
    var messages = mail.outgoingMessages().filter(function (m) {
      return m.id() === input.messageId;
    });
    if (messages.length !== 1 || messages[0].subject() !== input.markerSubject ||
        address(messages[0].sender()) !== address(input.expectedSender)) fail('identity');
  }
  stage('identity');
  checkMessage();
  var target = findMarkedComposer(process, input, checkMessage);
  var window = target.window, subject = target.subject, body = target.body;
  stage('focus');
  process.frontmost = true;
  window.attributes.byName('AXMain').value = true;
  for (var ready = 0; ready < 20; ready++) {
    if (process.frontmost() && attr(window, 'AXMain') === true) break;
    delay(0.1);
  }
  // Mail reports AXFocused=false on the editor even when it owns focus.
  // Click the exact AX element, then inspect the application's focused element.
  // AXPress on Mail's WebArea is a no-op on some macOS versions. Click inside
  // the verified editor's live bounds, then verify actual keyboard focus.
  // Empty composers can place a floating Siri writing control at the body's
  // top-left corner. Focus at the center; Cmd+Up below sets insertion to start.
  var position = attr(body, 'AXPosition'), size = attr(body, 'AXSize');
  if (!position || !size || size[0] < 20 || size[1] < 20 ||
      !process.frontmost() || attr(window, 'AXMain') !== true) fail('focus');
  var point = {x: position[0] + size[0] / 2, y: position[1] + size[1] / 2};
  var down = $.CGEventCreateMouseEvent(null, 1, point, 0);
  var up = $.CGEventCreateMouseEvent(null, 2, point, 0);
  $.CGEventPost(0, down);
  $.CGEventPost(0, up);
  function checkFocus() {
    checkMessage();
    // AX updates are asynchronous. Retry the read-only coherent snapshot, not
    // clicks, key events, clipboard writes, or insertion.
    for (var focusAttempt = 0; focusAttempt < 3; focusAttempt++) {
      var focused = attr(process, 'AXFocusedUIElement');
      var focusedWindow = focused ? attr(focused, 'AXWindow') : null;
      focusFailure = !process.frontmost() ? 'foreground'
        : attr(window, 'AXMain') !== true ? 'main-window'
        : !focused ? 'missing-element'
        : attr(focused, 'AXRole') !== 'AXWebArea' ? 'role'
        : attr(focused, 'AXDescription') !== 'message body' ? 'description'
        : !focusedWindow || attr(focusedWindow, 'AXTitle') !== input.markerSubject ? 'focused-window'
        : JSON.stringify(attr(focused, 'AXPosition')) !== JSON.stringify(attr(body, 'AXPosition')) ? 'position'
        : JSON.stringify(attr(focused, 'AXSize')) !== JSON.stringify(attr(body, 'AXSize')) ? 'size'
        : window.name() !== input.markerSubject ? 'window-title'
        : attr(subject, 'AXValue') !== input.markerSubject ? 'subject' : null;
      if (!focusFailure) return;
      if (focusAttempt < 2) delay(0.1);
    }
    console.log('NATIVE_FOCUS:' + focusFailure);
    fail('focus');
  }
  checkFocus();
  var files = Array.isArray(input.attachments) ? input.attachments : [];
  var prose = Boolean(normalized(input.body));
  if (!prose && !files.length) return;
  var beforeText = bodyText(body);
  var names = [], attachmentBaseline = null, archive = null;
  if (files.length) {
    stage('attachments');
    archive = webArchive(prose ? input.html : '', files);
    names = files.map(function (file) { return file.name; });
    // A forward's own attachments or a quoted file can share a pasted name.
    attachmentBaseline = attachmentButtons(body, names);
  }
  stage('clipboard-snapshot');
  var clipboard = $.NSPasteboard.generalPasteboard;
  var recovery = snapshotClipboard(clipboard, input.recoveryToken);
  var originalCount = recovery.originalCount;
  emit(recovery);
  var ownedCount = null;
  {
    stage('clipboard-write');
    var authored = $.NSPasteboardItem.alloc.init;
    // HTML and text stay as fallbacks; Mail prefers the archive when present.
    if ((prose && (!authored.setStringForType(input.html, $.NSPasteboardTypeHTML) ||
        !authored.setStringForType(input.body, $.NSPasteboardTypeString))) ||
        (archive && !authored.setDataForType(archive, 'com.apple.webarchive')) ||
        !authored.setStringForType('', 'org.nspasteboard.TransientType') ||
        !authored.setStringForType('', 'org.nspasteboard.AutoGeneratedType') ||
        !authored.setStringForType(input.recoveryToken, 'org.apple-mail-mcp.native-composition')) fail('clipboard');
    if (Number(clipboard.changeCount) !== originalCount) fail('clipboard');
    clipboard.clearContents;
    ownedCount = Number(clipboard.changeCount);
    emit({kind: 'clipboard-cleared', count: ownedCount});
    var wrote = writeItems(clipboard, [authored]);
    ownedCount = Number(clipboard.changeCount);
    emit({kind: 'clipboard-owned', count: ownedCount});
    if (!wrote) fail('clipboard');
    if (prose && (ObjC.unwrap(clipboard.stringForType($.NSPasteboardTypeHTML)) !== input.html ||
        ObjC.unwrap(clipboard.stringForType($.NSPasteboardTypeString)) !== input.body)) fail('clipboard');
    if (archive) {
      var written = clipboard.dataForType('com.apple.webarchive');
      if (written.isNil() || !written.isEqualToData(archive)) fail('clipboard');
    }
    stage('pre-cmd-up');
    checkFocus();
    events.keyCode(126, {using: ['command down']});
    stage('post-cmd-up');
    checkFocus();
    if (Number(clipboard.changeCount) !== ownedCount) fail('clipboard');
    stage('paste');
    events.keyCode(9, {using: ['command down']});
    // UI dispatch may finish before AX updates. Wait once, never paste twice.
    delay(files.length ? 1 : 0.3);
    stage('readback');
    checkFocus();
    if (Number(clipboard.changeCount) !== ownedCount) fail('clipboard');
    if (!verifiedInsertion(beforeText, bodyText(body), input.body, files.length)) fail('readback');
    if (files.length) {
      stage('attachment-arrival');
      waitForAttachments(body, names, attachmentBaseline);
    }
  }
  // The parent owns restoration, including after a timeout. A child cannot
  // safely recover from being killed halfway through its own restoration.
}
try {
  var data = $.NSFileHandle.fileHandleWithStandardInput.readDataToEndOfFile;
  var text = $.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding).js;
  main(JSON.parse(text));
  JSON.stringify({success: true});
} catch (error) {
  var allowed = ['accessibility', 'automation', 'automation-unavailable', 'identity', 'target', 'focus', 'clipboard', 'restore', 'readback', 'paste-unavailable', 'attachments', 'attachment-arrival'];
  JSON.stringify({success: false, code: allowed.indexOf(error.message) >= 0 ? error.message : 'bridge',
    stage: currentStage, focus: focusFailure || undefined, target: targetFailure || undefined,
    runtime: typeof error.errorNumber === 'number' ? error.errorNumber : undefined});
}
`;

const ERRORS: Record<string, string> = {
  accessibility:
    "Native Mail composition requires existing Accessibility permission for this runtime; no permission was requested.",
  automation:
    "Native Mail composition requires existing Automation permission for Mail and System Events; no permission was requested.",
  "automation-unavailable":
    "Native Mail composition could not contact Mail or start System Events; this does not establish a permission denial. Nothing was sent.",
  identity: "Native Mail composition refused an unexpected message, subject, or sender.",
  target:
    "Native Mail composition could not identify exactly one marked composer and message body.",
  focus: "Native Mail composition lost the verified composer or body focus; nothing was sent.",
  clipboard:
    "Native Mail composition stopped because the clipboard could not be preserved or changed unexpectedly.",
  restore: "Native Mail composition could not restore the previous clipboard; nothing was sent.",
  "paste-unavailable":
    "Mail's Paste command is unavailable for the verified editor; nothing was sent.",
  readback: "Native Mail composition could not verify the inserted prose; nothing was sent.",
  attachments:
    "Native Mail composition could not read or package an attachment for the editor; nothing was sent.",
  "attachment-arrival":
    "The pasted attachments did not appear in the composer; nothing was sent. Inspect the marked draft before retrying.",
  bridge:
    "Native Mail composition failed; nothing was sent. Inspect the marked draft and clipboard before retrying.",
};

function runNativeEditor(payload: Record<string, unknown>, timeout: number): NativeEditorResult {
  try {
    const result = spawnSync(
      "/usr/bin/osascript",
      ["-l", "JavaScript", "-e", NATIVE_EDITOR_SCRIPT],
      {
        input: JSON.stringify(payload),
        encoding: "utf8",
        timeout,
        killSignal: "SIGKILL",
        maxBuffer: MAX_NATIVE_EDITOR_OUTPUT_BYTES,
        stdio: ["pipe", "pipe", "pipe"],
      }
    );
    if (!recoverNativeClipboard(result.stdout ?? "", payload.recoveryToken)) {
      return { success: false, error: ERRORS.restore };
    }
    if (result.error || result.status !== 0) {
      const stages = [...(result.stderr ?? "").matchAll(/NATIVE_STAGE:([a-z-]+)/g)];
      const stage = stages.at(-1)?.[1];
      const timedOut = result.error && "code" in result.error && result.error.code === "ETIMEDOUT";
      return {
        success: false,
        error: ERRORS.bridge + diagnostic(stage, timedOut ? "timeout" : undefined),
      };
    }
    const response = JSON.parse(result.stdout.trim().split("\n").at(-1) ?? "");
    if (response.success === true) return { success: true };
    return {
      success: false,
      error:
        (ERRORS[response.code] ?? ERRORS.bridge) +
        diagnostic(response.stage, response.runtime, response.focus, response.target),
    };
  } catch {
    return { success: false, error: ERRORS.bridge };
  }
}

function diagnostic(stage: unknown, runtime: unknown, focus?: unknown, target?: unknown): string {
  const stages = [
    "startup",
    "preflight",
    "identity",
    "target-window",
    "target-body",
    "focus",
    "attachments",
    "clipboard-snapshot",
    "clipboard-write",
    "pre-cmd-up",
    "post-cmd-up",
    "paste",
    "readback",
    "attachment-arrival",
    "clipboard-restore",
  ];
  if (typeof stage !== "string" || !stages.includes(stage)) return "";
  const detail =
    runtime === "timeout" || (typeof runtime === "number" && Number.isInteger(runtime))
      ? `; runtime=${runtime}`
      : "";
  const focusDetail =
    typeof focus === "string" &&
    [
      "foreground",
      "main-window",
      "missing-element",
      "role",
      "description",
      "focused-window",
      "position",
      "size",
      "window-title",
      "subject",
    ].includes(focus)
      ? `; focus=${focus}`
      : "";
  const targetDetail =
    typeof target === "string" &&
    [
      "window-missing",
      "window-ambiguous",
      "subject-missing",
      "subject-ambiguous",
      "body-missing",
      "body-ambiguous",
    ].includes(target)
      ? `; target=${target}`
      : "";
  return ` [stage=${stage}${detail}${focusDetail}${targetDetail}]`;
}

export function nativeEditorPreflight(): NativeEditorResult {
  return runNativeEditor({ operation: "preflight" }, NATIVE_EDITOR_MIN_TIMEOUT_MS);
}

export function insertNativeAuthoredBody(input: NativeAuthoredBody): NativeEditorResult {
  if (
    !Number.isSafeInteger(input.messageId) ||
    input.messageId < 0 ||
    typeof input.markerSubject !== "string" ||
    !input.markerSubject ||
    typeof input.expectedSender !== "string" ||
    !input.expectedSender ||
    typeof input.body !== "string"
  ) {
    return { success: false, error: ERRORS.identity };
  }
  const paths = input.attachments ?? [];
  if (!paths.every((path) => typeof path === "string" && isAbsolute(path) && basename(path))) {
    return { success: false, error: ERRORS.attachments };
  }
  const body = input.body.normalize("NFC").replace(/\r\n?/g, "\n");
  return runNativeEditor(
    {
      ...input,
      body,
      attachments: paths.map((path) => ({ path, name: basename(path) })),
      recoveryToken: randomUUID(),
      operation: "insert",
      html: renderNativeAuthoredBody(body),
    },
    nativeEditorTimeoutMs(paths.length)
  );
}
