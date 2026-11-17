<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { VPptx } from '../src'
import { downloadBlob, pickFile, type OfficeApi } from '@voffice/shared'
const pptx = ref<OfficeApi | null>(null)
onMounted(() => { (window as unknown as { __pptx: unknown }).__pptx = pptx })
const toast = ref('')
function onSave(b: Blob, n: string) { downloadBlob(b, n) }
function onError(e: Error) { toast.value = e.message; setTimeout(() => toast.value = '', 6000) }
async function onOpenRequest() { const f = await pickFile('.pptx,.ppsx'); if (f[0]) pptx.value?.open(f[0]) }
async function onSaveRequest() { await pptx.value?.saveAs() }
</script>
<template>
  <div style="height:100%">
    <VPptx ref="pptx" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" />
    <div v-if="toast" style="position:fixed;left:50%;bottom:40px;transform:translateX(-50%);background:#ffffe1;border:1px solid #000;padding:8px 14px;z-index:5000">{{ toast }}</div>
  </div>
</template>
