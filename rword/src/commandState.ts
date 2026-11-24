/* The command states the original Quire's toolbars show.
 *
 * The original registers every command with an `enabled` predicate and, for the
 * toggle-like ones, a `checked` predicate (public/quire/js/commands.js). Its
 * `ui.refresh()` reads them on every state change and paints the result on the
 * button: `disabled` and `.dis` when the command cannot run, `.on` and
 * `aria-pressed` when it is toggled on (public/common/ui.js).
 *
 * This module carries the same predicates for the React editor. The state the
 * predicates read is gathered by the editor (index.tsx) and handed to
 * `commandEnabled` / `commandChecked`. */

export interface CommandState {
  /** the document is read-only */
  readOnly: boolean
  /** current document view */
  view: 'normal' | 'reading' | 'preview' | 'web' | 'print' | 'outline'
  /** the document can be edited: a document is open, it is not read-only and the
   *  view is not a reading-only one (the original's `canEdit`) */
  canEdit: boolean
  /** something is selected, and the selection is not collapsed (the original's `hasSel`) */
  hasSel: boolean
  /** the caret is in a table (the original's `inTable`) */
  inTable: boolean
  /** an object is selected (the original's `objSel`) */
  objSel: boolean
  /** a picture is selected (the original's `imgSel`) */
  imgSel: boolean
  /** a LettersArt object is selected (the original's `selectedObjIs(wordart)`) */
  wordartSel: boolean
  /** the header and footer are being edited (the original's `hfOnly`) */
  hfOnly: boolean
  /** the header and footer are linked to the previous section (the original's `A.hfLinked`) */
  hfLinked: boolean
  /** header and footer text is shown */
  hfShowText: boolean
  /** a mail-merge data source is loaded (the original's `hasData`) */
  hasData: boolean
  /** the merged data is being previewed (the original's `L.mailmerge.preview`) */
  mmPreview: boolean
  /** there is something to undo (the original's `D.canUndo`) */
  canUndo: boolean
  /** there is something to redo (the original's `D.canRedo`) */
  canRedo: boolean
  /** the run at the caret is bold (the original's `run().b`) */
  bold: boolean
  /** the run at the caret is italic (the original's `run().i`) */
  italic: boolean
  /** the run at the caret is underlined (the original's `run().u`) */
  underline: boolean
  /** the paragraph at the caret is aligned (the original's `para().jc`) */
  align: 'left' | 'center' | 'right' | 'justify'
  /** the paragraph at the caret is a bullet list (the original's `LS.info().fmt`) */
  listBullet: boolean
  /** the paragraph at the caret is a numbered list */
  listNumber: boolean
  /** the document tracks changes (the original's `doc().settings.track`) */
  trackChanges: boolean
  /** the formatting marks show (the original's `LY.opts.showMarks`) */
  showMarks: boolean
  /** the document map shows (the original's `A.opts.docMap`) */
  docMap: boolean
  /** the reviewing pane shows (the original's `!L.$('#revpane').hidden`) */
  reviewingPane: boolean
  /** the outline shows the first line only (the original's `L.outline.firstLine`) */
  outlineFirstLine: boolean
  /** the outline hides the text formatting (the original's `L.outline.plain`) */
  outlinePlain: boolean
  /** the table drawing tool (the original's `L.tables.tool`) */
  tableTool: string
  /** the object selection mode is on (the original's `A.selectObjectsMode`) */
  selectObjects: boolean
  /** the format painter is armed (the original's `A.painter.on`) */
  painterOn: boolean
  /** the highlighter is armed (the original's `A.highlighter.on`) */
  highlighterOn: boolean
  /** which toolbars show (the original's `A.tbShown`) */
  toolbars: Record<string, boolean>
}

/** The commands the original disables, and the rule for each.
 *
 * `canEdit` covers every command that changes the document; the rest need a
 * selection, a table, an object, the header and footer, or a data source. */
export function commandEnabled(id: string, s: CommandState): boolean {
  switch (id) {
    /* the document itself */
    case 'save':
    case 'saveAs':
    case 'saveWeb':
    case 'hfFormatPageNum':
      return true

    /* the undo history */
    case 'undo':
      return s.canUndo && !isReading(s)
    case 'redo':
      return s.canRedo && !isReading(s)

    /* cutting and copying */
    case 'cut':
      return s.canEdit && (s.hasSel || s.objSel)
    case 'copy':
      return s.hasSel || s.objSel
    case 'paste':
    case 'pasteSpecial':
    case 'pasteHyperlink':
    case 'clearFormats':
      return s.canEdit
    case 'clearContents':
      return s.canEdit && s.hasSel

    /* the view */
    case 'replace':
      return !isReading(s)
    case 'links':
      return false
    case 'editObject':
      return s.wordartSel
    case 'insertComment':
      return !isReading(s)

    /* inserting content */
    case 'insertClipArt':
    case 'insertPicture':
    case 'insertWordArt':
    case 'insertDiagram':
    case 'insertExcel':
    case 'textBox':
    case 'drawLine':
    case 'drawArrow':
    case 'drawRect':
    case 'drawOval':
    case 'hyperlink':
    case 'textDirection':
      return s.canEdit

    /* formatting */
    case 'bold':
    case 'italic':
    case 'underline':
    case 'alignLeft':
    case 'alignCenter':
    case 'alignRight':
    case 'justify':
    case 'numbering':
    case 'bullets':
    case 'incIndent':
    case 'decIndent':
    case 'borderBtn':
    case 'highlightBtn':
    case 'fontColorBtn':
    case 'painter':
    case 'sortAsc':
    case 'sortDesc':
    case 'promoteH1':
    case 'promote':
    case 'demote':
    case 'demoteBody':
    case 'moveUp':
    case 'moveDown':
    case 'updateTOC':
    case 'acceptChange':
    case 'rejectChange':
      return s.canEdit

    /* the reviewing toolbar */
    case 'trackChanges':
      return !s.readOnly
    case 'reviewingPane':
    case 'prevChange':
    case 'nextChange':
      return true

    /* the outlining toolbar: only in the outline view */
    case 'expand':
    case 'collapse':
    case 'olFirstLine':
    case 'olFormatting':
      return s.view === 'outline'
    case 'gotoTOC':
      return true

    /* tables */
    case 'drawTable':
    case 'eraser':
      return s.canEdit
    case 'mergeCells':
      return s.canEdit && s.inTable
    case 'splitCells':
    case 'tableAutoFormat':
    case 'distRows':
    case 'distCols':
    case 'autoSum':
      return s.canEdit && s.inTable

    /* the header and footer */
    case 'hfPageNum':
    case 'hfNumPages':
    case 'hfDate':
    case 'hfTime':
    case 'hfShowText':
    case 'hfSwitch':
    case 'hfPrev':
    case 'hfNext':
    case 'hfClose':
      return s.hfOnly
    case 'hfLinkPrev':
      return s.hfOnly
    case 'pageSetup':
      return true

    /* mail merge */
    case 'mmMainDoc':
    case 'mmOpenSource':
      return true
    case 'mmRecipients':
    case 'mmViewData':
    case 'mmFirst':
    case 'mmPrev':
    case 'mmNext':
    case 'mmLast':
    case 'mmMergeNew':
    case 'mmMergePrint':
      return s.hasData
    case 'mmAddress':
    case 'mmGreeting':
    case 'mmInsertField':
      return s.canEdit && s.hasData

    /* pictures and drawing objects */
    case 'picMoreContrast':
    case 'picLessContrast':
    case 'picMoreBright':
    case 'picLessBright':
    case 'picCrop':
    case 'picReset':
      return s.imgSel
    case 'picRotate':
    case 'fillColor':
    case 'lineColor':
      return s.canEdit && s.objSel
    case 'formatObject':
      return s.objSel
    case 'wordartEdit':
    case 'wordartGallery':
      return s.canEdit && s.wordartSel
    case 'selectObjects':
      return true

    default:
      return true
  }
}

/** The commands the original shows as toggled on, and the rule for each.
 *
 * `undefined` means the command has no `checked` predicate, so the original
 * leaves `aria-pressed` off it. */
export function commandChecked(id: string, s: CommandState): boolean | undefined {
  switch (id) {
    case 'printPreview':
      return s.view === 'preview'
    case 'painter':
      return s.painterOn
    case 'docMap':
      return s.docMap
    case 'showMarks':
      return s.showMarks
    case 'viewReading':
      return s.view === 'reading'
    case 'viewNormal':
      return s.view === 'normal'
    case 'viewWeb':
      return s.view === 'web'
    case 'viewPrint':
      return s.view === 'print'
    case 'viewOutline':
      return s.view === 'outline'
    case 'bold':
      return s.bold
    case 'italic':
      return s.italic
    case 'underline':
      return s.underline
    case 'alignLeft':
      return alignIs('left', s.align)
    case 'alignCenter':
      return alignIs('center', s.align)
    case 'alignRight':
      return alignIs('right', s.align)
    case 'justify':
      return alignIs('both', s.align)
    case 'numbering':
      return s.listNumber
    case 'bullets':
      return s.listBullet
    case 'highlightBtn':
      return s.highlighterOn
    case 'drawTable':
      return s.tableTool === 'draw'
    case 'eraser':
      return s.tableTool === 'erase'
    case 'trackChanges':
      return s.trackChanges
    case 'reviewingPane':
      return s.reviewingPane
    case 'olFirstLine':
      return s.outlineFirstLine
    case 'olFormatting':
      return !s.outlinePlain
    case 'hfShowText':
      return s.hfShowText
    case 'hfLinkPrev':
      return s.hfOnly && s.hfLinked
    case 'mmViewData':
      return s.mmPreview
    case 'selectObjects':
      return s.selectObjects
    default:
      if (id.startsWith('tb_')) return !!s.toolbars[id.slice(3)]
      return undefined
  }
}

/** The alignment test the original uses (public/quire/js/commands.js → alignIs):
 *  an unset alignment counts as left, and `both` also answers to `distribute`. */
function alignIs(jc: string, value: string): boolean {
  if (jc === 'left') return value === 'left' || value === 'start'
  if (jc === 'both') return value === 'both' || value === 'distribute'
  return value === jc
}

function isReading(s: CommandState): boolean {
  return s.view === 'reading' || s.view === 'preview'
}
