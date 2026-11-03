const { test } = require('node:test')
const assert = require('node:assert/strict')
const { readFileSync } = require('node:fs')
const vm = require('node:vm')
const { join } = require('node:path')

const source = readFileSync(join(__dirname, '../src/Code.js'), 'utf8')
function setup() {
  const writes = []
  const expansions = []
  const groups = []
  const formulas = []
  let groupPosition
  let released = false
  const range = {
    setValues(values) { writes.push(values); return this },
    setFontWeight() { return this }, setBackground() { return this },
    createFilter() {}, setNote(note) { this.note = note },
  }
  const sheet = {
    setName() {}, getSheetId() { return 42 }, getName() { return 'Folders 42' },
    getRange(row, column, count, columns) {
      return {
        ...range,
        setNote(note) { range.note = note },
        setFormulas(values) { formulas.push({ row, column, values }) },
        shiftRowGroupDepth(delta) { groups.push({ row, column, count, columns, delta }) },
      }
    },
    setRowGroupControlPosition(position) { groupPosition = position },
    setFrozenRows() {}, getMaxRows() { return 1000 },
    setColumnWidth() {}, setColumnWidths() {},
    insertRowsAfter(after, count) { expansions.push({ after, count }) },
  }
  const lock = { tryLock() { return true }, releaseLock() { released = true } }
  const context = vm.createContext({
    SpreadsheetApp: { GroupControlTogglePosition: { BEFORE: 'before' }, getActiveSpreadsheet() { return { insertSheet() { return sheet }, setActiveSheet() {} } } },
    LockService: { getDocumentLock() { return lock } },
    DriveApp: {},
  })
  vm.runInContext(source, context)
  return { context, writes, expansions, groups, formulas, groupPosition: () => groupPosition, range, lock, released: () => released }
}
function folder(name, children = []) {
  return {
    getName: () => name, getId: () => name, getUrl: () => `https://drive.google.com/drive/folders/${name}`,
    getSharingAccess: () => 'PRIVATE', getSharingPermission: () => 'NONE',
    getOwner: () => ({ getEmail: () => 'owner@example.org' }),
    getEditors: () => [{ getEmail: () => 'editor@example.org' }], getViewers: () => [],
    getFolders() {
      let i = 0
      return { hasNext: () => i < children.length, next: () => children[i++] }
    },
  }
}

test('extracts folder IDs and resource keys; rejects unrelated or malformed URLs', () => {
  const { context: c } = setup()
  assert.equal(c.parseFolderInput(' abc_123- ').id, 'abc_123-')
  assert.equal(c.parseFolderInput('https://drive.google.com/drive/u/0/folders/abc?resourcekey=key-1').resourceKey, 'key-1')
  assert.equal(c.parseFolderInput('https://drive.google.com/open?id=abc').id, 'abc')
  for (const value of ['', 'https://evil.org/folders/abc', 'https://drive.google.com/file/d/abc']) {
    assert.throws(() => c.parseFolderInput(value))
  }
})

test('requires a nonnegative whole depth', () => {
  const { context: c } = setup()
  assert.equal(c.parseDepth(' 0 '), 0)
  assert.equal(c.parseDepth(12), 12)
  for (const value of ['', -1, 1.5, 'abc', 'Infinity', '9007199254740992']) assert.throws(() => c.parseDepth(value))
})

test('walks depth first and includes only requested depths in a fresh report', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = () => folder('root', [folder('a', [folder('grandchild')]), folder('b')])
  const result = env.context.createFolderHierarchy('root', 1)
  const rows = env.writes[1]
  assert.equal(result.count, 3)
  assert.deepEqual(Array.from(rows, row => row[2]), ['root', 'root/a', 'root/b'])
  assert.equal(rows[1][0], '    a')
  assert.match(rows[0][4], /Restricted/)
  assert.equal(rows[0][7], 'editor@example.org')
  assert.equal(rows[0][9], '')
  assert.equal(env.released(), true)
})

test('depth zero never lists children and preserves resource keys', () => {
  const env = setup()
  const root = folder('root')
  root.getFolders = () => { throw new Error('Must not list children') }
  env.context.DriveApp.getFolderByIdAndResourceKey = (id, key) => {
    assert.equal(id, 'root'); assert.equal(key, 'secret'); return root
  }
  assert.equal(env.context.createFolderHierarchy('https://drive.google.com/drive/folders/root?resourcekey=secret', 0).count, 1)
})

test('sharing and child-listing errors are visible while other folders remain reportable', () => {
  const env = setup()
  const child = folder('child')
  child.getSharingAccess = () => { throw new Error('denied') }
  child.getFolders = () => { throw new Error('cannot list') }
  env.context.DriveApp.getFolderById = () => folder('root', [child, folder('sibling')])
  env.context.createFolderHierarchy('root', 2)
  assert.equal(env.writes[1].length, 3)
  assert.equal(env.writes[1][1][4], 'Unknown / unavailable')
  assert.match(env.writes[1][1][9], /denied/)
  assert.match(env.writes[1][1][9], /cannot list/)
})

test('escapes spreadsheet formulas in folder names', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = () => folder('=IMPORTXML("bad")')
  env.context.createFolderHierarchy('root', 0)
  assert.equal(env.writes[1][0][0], '\'=IMPORTXML("bad")')
})

test('releases the document lock when Drive access fails', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = () => { throw new Error('no access') }
  assert.throws(() => env.context.createFolderHierarchy('root', 1), /no access/)
  assert.equal(env.released(), true)
})

test('refuses concurrent runs without releasing another run’s lock', () => {
  const env = setup()
  env.lock.tryLock = () => false
  assert.throws(() => env.context.createFolderHierarchy('root', 1), /Another report/)
  assert.equal(env.released(), false)
})

test('marks capped reports partial and expands the sheet for all written rows', () => {
  const env = setup()
  const children = Array.from({ length: 3001 }, (_, i) => folder(`child${i}`))
  env.context.DriveApp.getFolderById = () => folder('root', children)
  const result = env.context.createFolderHierarchy('root', 1)
  assert.equal(result.count, 3000)
  assert.deepEqual(env.expansions, [{ after: 1000, count: 2001 }])
  assert.match(result.status, /PARTIAL/)
  assert.match(env.range.note, /PARTIAL/)
})

function mockUi(context, responses = []) {
  const prompts = []
  const alerts = []
  const menuItems = []
  let menuAdded = false
  const menu = {
    addItem(label, handler) { menuItems.push({ label, handler }); return this },
    addToUi() { menuAdded = true },
  }
  context.SpreadsheetApp.getUi = () => ({
    Button: { OK: 'ok', CANCEL: 'cancel' },
    ButtonSet: { OK_CANCEL: 'ok-cancel' },
    createMenu(label) { assert.equal(label, 'Folder hierarchy'); return menu },
    prompt(title, description, buttons) {
      prompts.push({ title, description, buttons })
      assert.ok(responses.length, 'Unexpected extra prompt')
      const response = responses.shift()
      return { getSelectedButton: () => response.button, getResponseText: () => response.text }
    },
    alert(message) { alerts.push(message) },
  })
  return { prompts, alerts, menuItems, menuAdded: () => menuAdded }
}

test('onOpen registers the report action without reading Drive', () => {
  const { context } = setup()
  const ui = mockUi(context)
  context.onOpen()
  assert.deepEqual(ui.menuItems, [{ label: 'Create folder report…', handler: 'promptFolderHierarchy' }])
  assert.equal(ui.menuAdded(), true)
})

for (const cancelAt of [0, 1]) {
  test(`canceling prompt ${cancelAt + 1} creates no report`, () => {
    const env = setup()
    const responses = [{ button: 'ok', text: 'root' }, { button: 'ok', text: '1' }]
    responses[cancelAt].button = 'cancel'
    const ui = mockUi(env.context, responses)
    env.context.promptFolderHierarchy()
    assert.equal(ui.prompts.length, cancelAt + 1)
    assert.equal(env.writes.length, 0)
    assert.equal(ui.alerts.length, 0)
  })
}

test('menu prompts invoke the real report function and show its result', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = id => {
    assert.equal(id, 'root')
    return folder('root', [folder('child')])
  }
  const ui = mockUi(env.context, [{ button: 'ok', text: 'root' }, { button: 'ok', text: '1' }])
  env.context.promptFolderHierarchy()
  assert.deepEqual(ui.prompts.map(prompt => prompt.title), [
    'Step 1 of 2: Starting folder', 'Step 2 of 2: Recursion depth',
  ])
  assert.equal(env.writes[1].length, 2)
  assert.equal(ui.alerts.length, 1)
  assert.match(ui.alerts[0], /Report created: Folders 42\n2 folders/)
})

test('menu displays input errors without attempting a report', () => {
  const env = setup()
  const ui = mockUi(env.context, [{ button: 'ok', text: 'root' }, { button: 'ok', text: '-1' }])
  env.context.promptFolderHierarchy()
  assert.equal(env.writes.length, 0)
  assert.match(ui.alerts[0], /Could not create report: Depth must be/)
})

test('sharing labels distinguish public, link, organization, and restricted access', () => {
  const { context } = setup()
  for (const [access, label] of [
    ['ANYONE', 'Public on the web'],
    ['ANYONE_WITH_LINK', 'Anyone with the link'],
    ['DOMAIN', 'Organization'],
    ['DOMAIN_WITH_LINK', 'Organization with the link'],
    ['PRIVATE', 'Restricted (specific accounts may have access)'],
  ]) {
    const item = folder('root')
    item.getSharingAccess = () => access
    assert.equal(context.sharingDetails(item)[0], label)
  }
})

test('absent owners and empty account lists are blank; hidden emails remain explicit', () => {
  const { context } = setup()
  const item = folder('root')
  item.getOwner = () => null
  item.getEditors = () => [{ getEmail: () => '' }]
  const details = context.sharingDetails(item)
  assert.equal(details[2], '')
  assert.equal(details[3], '(email hidden)')
  assert.equal(details[4], '')
  assert.equal(details[5], '')
})

test('an iterator failure records the affected parent and continues with its sibling', () => {
  const env = setup()
  const broken = folder('broken')
  broken.getFolders = () => ({ hasNext() { throw new Error('iterator denied') } })
  env.context.DriveApp.getFolderById = () => folder('root', [broken, folder('sibling')])
  env.context.createFolderHierarchy('root', 2)
  assert.deepEqual(Array.from(env.writes[1], row => row[2]), ['root', 'root/broken', 'root/sibling'])
  assert.match(env.writes[1][1][9], /Child listing incomplete: iterator denied/)
})

test('the time budget produces a partial report without waiting in real time', () => {
  const env = setup()
  let now = 0
  vm.runInContext('Date.now = () => 0', env.context)
  const root = folder('root', [folder('child')])
  root.getFolders = () => {
    now += 240001
    env.context.now = now
    vm.runInContext('Date.now = () => now', env.context)
    return { hasNext() { throw new Error('Should stop before listing') } }
  }
  env.context.DriveApp.getFolderById = () => root
  const result = env.context.createFolderHierarchy('root', 1)
  assert.equal(result.count, 1)
  assert.match(result.status, /PARTIAL/)
  assert.equal(env.released(), true)
})


test('a failed depth prompt displays the error instead of silently stopping', () => {
  const env = setup()
  const ui = mockUi(env.context, [{ button: 'ok', text: 'root' }])
  const getUi = env.context.SpreadsheetApp.getUi
  env.context.SpreadsheetApp.getUi = () => {
    const service = getUi()
    const prompt = service.prompt
    service.prompt = (...args) => {
      if (args[0] === 'Step 2 of 2: Recursion depth') throw new Error('Depth dialog failed')
      return prompt(...args)
    }
    return service
  }
  env.context.promptFolderHierarchy()
  assert.deepEqual(ui.alerts, ['Could not create report: Depth dialog failed'])
  assert.equal(env.writes.length, 0)
})


test('groups descendants at each hierarchy level while leaving their parents outside each group', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = () => folder('root', [
    folder('a', [folder('a1', [folder('a1-child')]), folder('a2')]),
    folder('b', [folder('b1')]), folder('leaf'),
  ])
  env.context.createFolderHierarchy('root', 3)
  assert.equal(env.groupPosition(), 'before')
  assert.deepEqual(env.groups.map(({ row, count, delta }) => ({ row, count, delta })), [
    { row: 3, count: 7, delta: 1 }, // all root descendants
    { row: 4, count: 3, delta: 1 }, // a's descendants
    { row: 8, count: 1, delta: 1 }, // b's descendants
    { row: 5, count: 1, delta: 1 }, // a1's descendants
  ])
  const depths = Array(10).fill(0)
  for (const group of env.groups) {
    for (let row = group.row; row < group.row + group.count; row++) depths[row] += group.delta
  }
  assert.deepEqual(depths.slice(2), [0, 1, 2, 3, 2, 1, 2, 1])
})

test('depth-zero reports have no groups; deeper reports respect the eight-level Sheets limit', () => {
  const env = setup()
  env.context.DriveApp.getFolderById = () => folder('root')
  env.context.createFolderHierarchy('root', 0)
  assert.deepEqual(env.groups, [])
  const rows = Array.from({ length: 11 }, (_, depth) => ['folder', depth])
  env.context.groupFolderRows(env.context.SpreadsheetApp.getActiveSpreadsheet().insertSheet(), rows)
  assert.equal(env.groups.length, 8)
  assert.ok(env.groups.every(group => group.row >= 3))
})

test('writes ten columns without IDs and writes compact hyperlink formulas separately', () => {
  const env = setup()
  const root = folder('root')
  root.getId = () => { throw new Error('Folder IDs must not be read for the report') }
  root.getUrl = () => 'https://drive.google.com/drive/folders/root?resourcekey=key'
  env.context.DriveApp.getFolderById = () => root
  env.context.createFolderHierarchy('root', 0)
  const headers = env.writes[0][0]
  assert.equal(headers.length, 10)
  assert.equal(headers.includes('Folder ID'), false)
  assert.equal(headers[3], 'Folder URL')
  assert.equal(env.writes[1][0].length, 10)
  assert.equal(env.writes[1][0][3], '')
  assert.equal(env.formulas[0].column, 4)
  assert.equal(env.formulas[0].values[0][0], '=HYPERLINK("https://drive.google.com/drive/folders/root?resourcekey=key", "link")')
})

test('escapes double quotes in hyperlink URLs', () => {
  const env = setup()
  const root = folder('root')
  root.getUrl = () => 'https://drive.google.com/drive/folders/a"b'
  env.context.DriveApp.getFolderById = () => root
  env.context.createFolderHierarchy('root', 0)
  assert.equal(env.formulas[0].values[0][0], '=HYPERLINK("https://drive.google.com/drive/folders/a""b", "link")')
})

test('EXTERNAL_EMAIL returns true for non-SCELC email addresses and false for SCELC addresses', () => {
  const env = setup()
  const EXTERNAL_EMAIL = env.context.EXTERNAL_EMAIL
  assert.equal(EXTERNAL_EMAIL('user@cca.edu'), true)
  assert.equal(EXTERNAL_EMAIL('user@scelc.org, user@cca.edu'), true)
  assert.equal(EXTERNAL_EMAIL('user@scelc.org, user@cca.edu, user@scelc.org'), true)
  assert.equal(EXTERNAL_EMAIL('user@cca.edu, user@scelc.org'), true)
  assert.equal(EXTERNAL_EMAIL('user@cca.edu, user@scelc.org '), true) // trailing space
  assert.equal(EXTERNAL_EMAIL('user@scelc.org'), false)
  assert.equal(EXTERNAL_EMAIL('user@scelc.org, phette@scelc.org'), false)
})
