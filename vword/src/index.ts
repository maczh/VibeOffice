import type { App } from 'vue'
import '@voffice/shared/styles'
import './styles/app.css'
import VWord from './VWord.vue'

export { VWord }
export type { OfficeApi, OfficeProps, OfficeEmits } from '@voffice/shared'
export default { install(app: App) { app.component('VWord', VWord) } }
