import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import RWord from './index'
/* The classic chrome stylesheets — the React UI renders the same DOM as public/quire,
 * so loading these makes the two interfaces render identically. */
import '../../public/common/looks.css'
import '../../public/common/ui.css'
import '../../public/fonts/fonts.css'
import '../../public/common/scout/scout.css'
import './quire.css'
import { initLanguage } from './i18n'

initLanguage()
document.documentElement.dataset.look = 'classic'
document.documentElement.dataset.layout = 'bars'
document.documentElement.dataset.app = 'quire'
document.documentElement.dataset.appName = 'Quire'
document.title = 'Quire'
createRoot(document.getElementById('root')!).render(<StrictMode><RWord /></StrictMode>)

/* Scout is the suite's shared assistant; load the same script the original pages use. */
// @ts-expect-error The shared assistant is a classic script without type declarations.
void import('../../public/common/scout/scout.js').then(() => {
  const host = window as unknown as { L?: { Scout?: new (options: Record<string, unknown>) => { show: () => void; placeAt: (x: number, y: number) => void } } }
  const Scout = host.L?.Scout
  if (!Scout) return
  const size = 154
  const height = Math.round(size * 235 / 260)
  const scout = new Scout({ size, balloon: false, idle: false, label: 'Scout, the assistant' })
  scout.placeAt(Math.max(4, window.innerWidth - size - 24), Math.max(4, window.innerHeight - height - 56))
  scout.show()
})
