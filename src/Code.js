/** @OnlyCurrentDoc */

/**
 * Detect non-SCELC emails in a comma-separated list
 * @param {any} str - comma-separated list of emails
 * @returns {boolean} - true if any email does not end with @scelc.org
 */
function EXTERNAL_EMAIL(str) {
  const emails = str.split(", ")
  return emails.some(email => !email.trim().endsWith("@scelc.org"))
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Folder hierarchy')
    .addItem('Create folder report…', 'promptFolderHierarchy').addToUi()
}

function promptFolderHierarchy() {
  const ui = SpreadsheetApp.getUi()
  try {
    const folder = ui.prompt('Step 1 of 2: Starting folder', 'Enter a Google Drive folder ID or URL, then click OK to choose the recursion depth.', ui.ButtonSet.OK_CANCEL)
    if (folder.getSelectedButton() !== ui.Button.OK) return
    const depth = ui.prompt('Step 2 of 2: Recursion depth', 'Enter a whole number: 0 = starting folder only; 1 = immediate children; 2 = include grandchildren, etc.', ui.ButtonSet.OK_CANCEL)
    if (depth.getSelectedButton() !== ui.Button.OK) return
    const result = createFolderHierarchy(folder.getResponseText(), depth.getResponseText())
    ui.alert(`Report created: ${result.sheetName}\n${result.count} folders.\n${result.status}`)
  } catch (error) {
    ui.alert(`Could not create report: ${error.message}`)
  }
}

function parseFolderInput(input) {
  const value = String(input).trim()
  if (/^[\w-]+$/.test(value)) return { id: value, resourceKey: '' }
  if (!/^https:\/\/drive\.google\.com\//i.test(value)) throw new Error('Enter a folder ID or an https://drive.google.com/ URL.')
  const match = value.match(/\/folders\/([\w-]+)/) || value.match(/[?&]id=([\w-]+)/)
  if (!match) throw new Error('The URL does not contain a folder ID.')
  const key = value.match(/[?&]resourcekey=([\w-]+)/i)
  return { id: match[1], resourceKey: key ? key[1] : '' }
}

function parseDepth(input) {
  const value = String(input).trim()
  if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value))) {
    throw new Error('Depth must be a nonnegative whole number.')
  }
  return Number(value)
}

function safeCell(value) {
  const text = String(value ?? '')
  return /^[=+\-@\t\r\n]/.test(text) ? `'${text}` : text
}

function sharingDetails(folder) {
  const notes = []
  function read(label, getter) {
    try { return getter() } catch (error) {
      notes.push(`${label}: ${error.message}`)
      return 'Unknown / unavailable'
    }
  }
  const access = read('General access', () => String(folder.getSharingAccess()))
  const labels = {
    ANYONE: 'Public on the web',
    ANYONE_WITH_LINK: 'Anyone with the link',
    DOMAIN: 'Organization',
    DOMAIN_WITH_LINK: 'Organization with the link',
    PRIVATE: 'Restricted (specific accounts may have access)',
  }
  const accounts = users => users.map(user => user.getEmail() || '(email hidden)').join(', ')
  return [
    labels[access] || access,
    read('General role', () => String(folder.getSharingPermission())),
    read('Owner', () => { const owner = folder.getOwner(); return owner ? owner.getEmail() || '(email hidden)' : '' }),
    read('Editors', () => accounts(folder.getEditors())),
    read('Viewers/commenters', () => accounts(folder.getViewers())),
    notes.join(' '),
  ]
}

/** Creates a new tab in the bound spreadsheet. Root is depth zero. */
function createFolderHierarchy(startingFolder, depthLimit) {
  const parsed = parseFolderInput(startingFolder)
  const maxDepth = parseDepth(depthLimit)
  const lock = LockService.getDocumentLock()
  if (!lock.tryLock(1000)) throw new Error('Another report is running. Try again when it finishes.')
  try {
    const root = parsed.resourceKey
      ? DriveApp.getFolderByIdAndResourceKey(parsed.id, parsed.resourceKey)
      : DriveApp.getFolderById(parsed.id)
    const book = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = book.insertSheet()
    sheet.setName(`${root.getName()} Hierarchy`)
    const headers = ['Folder', 'Depth', 'Path', 'Folder URL', 'General access', 'General role', 'Owner', 'Editors', 'Viewers/commenters', 'Notes']
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3')
    sheet.setFrozenRows(1)
    // Leave time for writing before Apps Script's execution limit.
    const deadline = Date.now() + 240000
    const rows = []
    const stack = [{ folder: root, depth: 0, parent: '' }]
    let status = 'Complete within requested depth.'
    while (stack.length) {
      if (Date.now() >= deadline || rows.length >= 3000) {
        status = 'PARTIAL: time or 3,000-folder limit reached. Run smaller subtrees to cover the remainder.'
        break
      }
      const item = stack.pop()
      if (item.iterator) {
        try {
          if (item.iterator.hasNext()) {
            const child = item.iterator.next()
            stack.push(item, { folder: child, depth: item.depth, parent: item.parent })
          }
        } catch (error) {
          rows[item.row][9] += ` Child listing incomplete: ${error.message}`
        }
        continue
      }
      const name = item.folder.getName()
      const path = item.parent ? `${item.parent}/${name}` : name
      rows.push([
        `${'    '.repeat(Math.min(item.depth, 100))}${name}`, item.depth, path,
        item.folder.getUrl(), ...sharingDetails(item.folder),
      ])
      if (item.depth < maxDepth) {
        try {
          stack.push({ iterator: item.folder.getFolders(), depth: item.depth + 1, parent: path, row: rows.length - 1 })
        } catch (error) {
          rows[rows.length - 1][9] += ` Child listing incomplete: ${error.message}`
        }
      }
    }
    const required = rows.length + 1
    if (sheet.getMaxRows() < required) sheet.insertRowsAfter(sheet.getMaxRows(), required - sheet.getMaxRows())
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows.map(row => row.map((cell, column) => column === 3 ? '' : typeof cell === 'number' ? cell : safeCell(cell))))
      // Write only our generated formulas as formulas; folder names remain escaped text.
      sheet.getRange(2, 4, rows.length, 1).setFormulas(rows.map(row => [
        `=HYPERLINK("${row[3].replace(/"/g, '""')}", "link")`,
      ]))
      groupFolderRows(sheet, rows)
      sheet.getRange(1, 1, required, headers.length).createFilter()
    }
    sheet.setColumnWidth(1, 320)
    sheet.setColumnWidth(2, 65)
    sheet.setColumnWidths(3, 8, 240)
    sheet.setColumnWidth(4, 65)
    sheet.getRange(1, 1).setNote(`${status}\nGenerated: ${new Date().toISOString()}\nDepth limit: ${maxDepth}\nChild order follows Drive. Read Notes for per-folder failures.`)
    book.setActiveSheet(sheet)
    return { sheetName: sheet.getName(), count: rows.length, status }
  } finally {
    lock.releaseLock()
  }
}


/** Group contiguous descendants at each supported outline level; parents stay visible. */
function groupFolderRows(sheet, rows) {
  sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE)
  for (let level = 1; level <= 8; level++) {
    let start = -1
    for (let index = 0; index <= rows.length; index++) {
      const inGroup = index < rows.length && rows[index][1] >= level
      if (inGroup && start === -1) start = index
      if (!inGroup && start !== -1) {
        // Data starts at sheet row 2, after the header.
        sheet.getRange(start + 2, 1, index - start, 1).shiftRowGroupDepth(1)
        start = -1
      }
    }
  }
}
