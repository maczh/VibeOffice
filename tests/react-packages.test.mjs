import assert from 'node:assert/strict'
import { test } from 'node:test'

test('React package public exports stay available', async () => {
  const packages = [
    ['@voffice/rword', 'RWord'],
    ['@voffice/rexcel', 'RExcel'],
    ['@voffice/rpptx', 'RPptx'],
  ]
  for (const [name, component] of packages) {
    const main = await import(name)
    const engine = await import(`${name}/engine`)
    assert.equal(typeof main[component], 'object')
    assert.ok(Object.keys(engine).length > 0)
  }
})

test('React shared online contract is exported', async () => {
  const shared = await import('@voffice/react-shared')
  assert.equal(typeof shared.fetchOnline, 'function')
  assert.equal(typeof shared.saveOnline, 'function')
  assert.equal(typeof shared.downloadBlob, 'function')
  assert.equal(typeof shared.ErrorBoundary, 'function')
})

test('React shared common migration exports are public', async () => {
  const shared = await import('@voffice/react-shared')
  for (const name of ['StartCenter', 'LanguageSelect', 'PwaUpdate', 'Assistant', 'FileDropZone', 'CommandSearch', 'OnlineSourceDialog']) assert.equal(typeof shared[name], 'function', name)
  for (const name of ['useRecentFiles', 'useLook', 'useI18n', 'usePwa', 'useAssistant', 'useDirtyGuard']) assert.equal(typeof shared[name], 'function', name)
  assert.equal(typeof shared.commandsFromMenus, 'function')
})

test('React shared UI exports cover the migrated shared component set', async () => {
  const shared = await import('@voffice/react-shared')
  const pairs = {
    VoCommandSearch: 'CommandSearch', VoContextMenu: 'ContextMenu', VoStartCenter: 'StartCenter',
    VoLanguageSelect: 'LanguageSelect', VoPwaUpdate: 'PwaUpdate', VoAssistant: 'Assistant',
    VoFileDropZone: 'FileDropZone', VoOnlineDialog: 'OnlineSourceDialog',
  }
  for (const reactName of Object.values(pairs)) assert.equal(typeof shared[reactName], 'function', reactName)
})
