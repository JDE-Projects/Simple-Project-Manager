// ui_drive scenarios for Simple Project Manager. Each begins from a recovery
// workbook supplied by its fixture. A scenario that clicks Save restores it
// first, so Save uses its existing filename and never opens a native file
// dialog.

const RECOVERY_WARNING = "Couldn't delete the recovery copy, so it may be offered again next launch.";

function toastShows(helpers, text) {
  return helpers.waitFor(
    "document.getElementById('toast').classList.contains('show') && " +
    `document.getElementById('toast').textContent === ${JSON.stringify(text)}`, 10000
  ).then(() => true).catch(() => false);
}

function toastText(helpers) {
  return helpers.evaluate("document.getElementById('toast').textContent");
}

async function restoreFixture(helpers) {
  const { click, evaluate, waitFor, check } = helpers;

  await waitFor("window.pywebview && pywebview.api && typeof pywebview.api.get_state === 'function'", 10000);
  const state = await evaluate("pywebview.api.get_state()");
  check("page and bridge answer", !!state && typeof state.version === "string", JSON.stringify(state));

  const promptShown = await waitFor(
    "document.getElementById('modal-scrim').classList.contains('open')", 10000
  ).then(() => true).catch(() => false);
  check("restore prompt appears", promptShown);
  if (promptShown) await click("#modal-ok");
  const restored = await waitFor(
    "STATE.items.some(item => item.name === 'Fixture phase') && " +
    "STATE.items.some(item => item.name === 'Fixture step') && " +
    "document.getElementById('tree').textContent.includes('Fixture phase') && " +
    "document.getElementById('tree').textContent.includes('Fixture step')", 10000
  ).then(() => true).catch(() => false);
  check("restored fixture items show", restored, await evaluate("JSON.stringify(STATE.items)"));

  return state;
}

async function checkVersion(helpers, state) {
  const { evaluate, check } = helpers;
  const label = await evaluate("document.getElementById('verLabel').textContent");
  check("version label equals APP_VERSION from the bridge", label === "v" + state.version,
    `${label} vs v${state.version}`);
}

async function checkThemes(helpers) {
  const { click, evaluate, waitFor, check, screenshot } = helpers;
  const startLight = await evaluate("document.body.classList.contains('light')");
  await click("#theme-btn");
  await waitFor(`document.body.classList.contains('light') === ${!startLight}`, 5000);
  const otherLight = await evaluate("document.body.classList.contains('light')");
  check("theme toggles to the other theme", otherLight === !startLight, String(otherLight));
  await screenshot(otherLight ? "theme-light" : "theme-dark");
  await click("#theme-btn");
  await waitFor(`document.body.classList.contains('light') === ${startLight}`, 5000);
  const restoredLight = await evaluate("document.body.classList.contains('light')");
  check("theme toggles back", restoredLight === startLight, String(restoredLight));
  await screenshot(restoredLight ? "theme-light" : "theme-dark");
}

async function smoke(helpers) {
  const { click, type, evaluate, waitFor, check } = helpers;
  const state = await restoreFixture(helpers);
  await checkVersion(helpers, state);
  await checkThemes(helpers);

  await click(".toolbar button[onclick=\"addPhase()\"]");
  // The edit panel slides in; wait until the name field is no longer covered.
  await waitFor(
    "(() => { const f = document.getElementById('f_name'); if (!f || " +
    "!document.getElementById('panel').classList.contains('open')) return false; " +
    "const r = f.getBoundingClientRect(); " +
    "return document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2) === f; })()",
    5000
  );
  await type("#f_name", "Added smoke phase");
  await click("#panel .panel-foot .primary");
  const added = await waitFor(
    "STATE.items.some(item => item.type === 'Phase' && item.name === 'Added smoke phase')", 10000
  ).then(() => true).catch(() => false);
  check("a phase was added and renamed through the UI", added,
    await evaluate("JSON.stringify(STATE.items)"));

  await click(".toolbar .btn.primary");
  const saved = await waitFor(
    "document.getElementById('toast').classList.contains('show') && " +
    "document.getElementById('toast').textContent === 'Saved' && STATE.dirty === false", 10000
  ).then(() => true).catch(() => false);
  check("normal Save succeeds and clears the unsaved state", saved,
    await evaluate("JSON.stringify({toast: document.getElementById('toast').textContent, dirty: STATE.dirty})"));
}

async function saveError(helpers) {
  const { click, evaluate, waitFor, check, screenshot } = helpers;
  await restoreFixture(helpers);
  await click(".toolbar .btn.primary");
  const toastShown = await waitFor("document.getElementById('toast').classList.contains('show')", 10000)
    .then(() => true).catch(() => false);
  const toast = await evaluate("document.getElementById('toast').textContent");
  const expected = ["Couldn't save — is the file open in Excel?", "Couldn't save the project file."];
  check("failed direct Save shows a plain-language message", toastShown && expected.includes(toast) &&
    !/exception|traceback|errno|winerror/i.test(toast), toast);
  check("project remains marked unsaved after failed Save", await evaluate("STATE.dirty === true"),
    await evaluate("JSON.stringify({dirty: STATE.dirty, meta: document.getElementById('meta').textContent})"));
  await screenshot("save-error-toast");
}

async function recoveryLockedSave(helpers) {
  const { click, evaluate, check, screenshot } = helpers;
  await restoreFixture(helpers);

  await click(".toolbar .btn.primary");
  check("Save reports the undeletable recovery copy",
    await toastShows(helpers, `Saved. ${RECOVERY_WARNING}`), await toastText(helpers));
  check("the save itself still counts", await evaluate("STATE.dirty === false"));
  await screenshot("recovery-locked-save-toast");

  await click(".toolbar button[onclick=\"newProject()\"]");
  check("New reports the undeletable recovery copy",
    await toastShows(helpers, `New project. ${RECOVERY_WARNING}`), await toastText(helpers));
}

async function recoveryLockedDecline(helpers) {
  const { click, evaluate, waitFor, check, screenshot } = helpers;
  await waitFor("document.getElementById('modal-scrim').classList.contains('open')", 10000);
  await click("#modal-cancel");

  check("declining reports the undeletable recovery copy",
    await toastShows(helpers, RECOVERY_WARNING), await toastText(helpers));
  await screenshot("recovery-locked-decline-toast");
  check("declining leaves no file name and no restored items",
    await evaluate("STATE.filename === '' && !STATE.items.some(item => item.name === 'Fixture phase')"),
    await evaluate("JSON.stringify({filename: STATE.filename, items: STATE.items})"));

  // The ordinary message time is 2.6 seconds, and a warning stays up for 6.
  await new Promise(resolve => setTimeout(resolve, 3000));
  check("the warning outlasts the ordinary message time",
    await evaluate("document.getElementById('toast').classList.contains('show')"));
}

async function recoveryLockedUnreadable(helpers) {
  const { evaluate, waitFor, check } = helpers;
  // The message may have faded by the time the driver connects, so this reads
  // its text, which stays in place after it fades.
  const reported = await waitFor(
    `document.getElementById('toast').textContent === ${JSON.stringify(RECOVERY_WARNING)}`, 10000
  ).then(() => true).catch(() => false);
  check("an unreadable, undeletable recovery copy is reported", reported, await toastText(helpers));
  check("no restore is offered for an unreadable copy",
    await evaluate("!document.getElementById('modal-scrim').classList.contains('open')"));
}

export default async function scenario(helpers) {
  const scenarios = {
    "save-error": saveError,
    "recovery-locked-save": recoveryLockedSave,
    "recovery-locked-decline": recoveryLockedDecline,
    "recovery-locked-unreadable": recoveryLockedUnreadable,
  };
  const chosen = helpers.fixture && scenarios[helpers.fixture.mode];
  return (chosen || smoke)(helpers);
}
