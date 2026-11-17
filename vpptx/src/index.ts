import type { App } from 'vue'
import '@voffice/shared/styles'
import './styles/app.css'
import VPptx from './VPptx.vue'
export { VPptx }
export type { OfficeApi, OfficeProps, OfficeEmits } from '@voffice/shared'
export default { install(app: App) { app.component('VPptx', VPptx) } }
