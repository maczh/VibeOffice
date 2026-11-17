<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { VWord } from '../src'
import { downloadBlob, pickFile, type OfficeApi } from '@voffice/shared'
const word = ref<OfficeApi | null>(null)
onMounted(() => { (window as unknown as { __word: unknown }).__word = word })
const toast = ref('')
function onSave(blob: Blob, name: string) { downloadBlob(blob, name) }
function onError(e: Error) { toast.value = e.message; setTimeout(() => toast.value = '', 6000) }
async function onOpenRequest() { const f = await pickFile('.docx,.docm,.dotx,.dotm'); if (f[0]) word.value?.open(f[0]) }
async function onSaveRequest() { await word.value?.saveAs() }
</script>
<template>
  <div style="height:100%">
    <VWord ref="word" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" />
    <div v-if="toast" style="position:fixed;left:50%;bottom:40px;transform:translateX(-50%);background:#ffffe1;border:1px solid #000;padding:8px 14px;z-index:5000">{{ toast }}</div>
  </div>
</template>
