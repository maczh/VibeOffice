/** React migration map for the reference Quire scripts under public/quire/js. */
export const quireScriptMigration = {
  migratedServices: ['find.js', 'markdown.js', 'mdio.js', 'htmlio.js', 'autocorrect.js', 'mailmerge.js', 'tables.js', 'fields.js', 'spell.js', 'drawing.js', 'drawtool.js', 'rtf.js', 'review.js', 'preserve.js', 'preserve-drawing.js'],
  migratedUi: ['app.js', 'dialogs.js', 'dialogs2.js', 'dialogs3.js', 'ribbon.js', 'rulers.js', 'panes.js', 'editor.js', 'outline.js'],
  engineModules: ['commands.js', 'dmodel.js', 'docx-read.js', 'docx-write.js', 'editor.js', 'layout.js', 'render.js', 'ops.js', 'portions.js'],
  remainingModules: ['autocorrect.js', 'drawing.js', 'drawtool.js', 'fields.js', 'mailmerge.js', 'omml.js', 'ops.js', 'outline.js', 'panes.js', 'portions.js', 'preserve-drawing.js', 'preserve.js', 'review.js', 'rtf.js', 'spell.js', 'tables.js', 'templates.js'],
} as const
