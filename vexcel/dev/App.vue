<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { VExcel } from '../src'
import { downloadBlob, pickFile, type OfficeApi } from '@voffice/shared'
const excel = ref<OfficeApi | null>(null)
onMounted(() => { (window as unknown as { __excel: unknown }).__excel = excel })
const toast = ref('')
function onSave(b: Blob, n: string) { downloadBlob(b, n) }
function onError(e: Error) { toast.value = e.message; setTimeout(() => toast.value = '', 6000) }
async function onOpenRequest() { const f = await pickFile('.xlsx,.xlsm,.csv'); if (f[0]) excel.value?.open(f[0]) }
async function onSaveRequest() { await excel.value?.saveAs() }
</script>
<template>
  <div style="height:100%">
    <VExcel ref="excel" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" />
    <div v-if="toast" style="position:fixed;left:50%;bottom:40px;transform:translateX(-50%);background:#ffffe1;border:1px solid #000;padding:8px 14px;z-index:5000">{{ toast }}</div>
  </div>
</template>
