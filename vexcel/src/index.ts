import type { App } from 'vue'
import '@voffice/shared/styles'
import './styles/app.css'
import VExcel from './VExcel.vue'
export { VExcel }
export type { OfficeApi, OfficeProps, OfficeEmits } from '@voffice/shared'
export default { install(app: App) { app.component('VExcel', VExcel) } }
