<script setup lang="ts">
/* The integrated run: all three office components on one page, each in its own tab.
 * They stay mounted (v-show) so switching tabs keeps every document's state. */
import { ref, computed } from 'vue'
import { VWord } from '@voffice/vword'
import { VExcel } from '@voffice/vexcel'
import { VPptx } from '@voffice/vpptx'
import { downloadBlob, pickFile, useLook, type Look, type OfficeApi } from '@voffice/shared'

const { look } = useLook()
const tab = ref<'word' | 'excel' | 'pptx'>('word')
const word = ref<OfficeApi | null>(null)
const excel = ref<OfficeApi | null>(null)
const pptx = ref<OfficeApi | null>(null)
const toast = ref('')
let toastTimer: ReturnType<typeof setTimeout> | undefined

const active = computed(() => (tab.value === 'word' ? word.value : tab.value === 'excel' ? excel.value : pptx.value))
const accept = computed(() => tab.value === 'word' ? '.docx,.docm,.dotx,.dotm' : tab.value === 'excel' ? '.xlsx,.xlsm,.xltx,.xltm,.csv,.txt' : '.pptx,.pptm,.ppsx,.ppsm,.potx,.potm')

function notify(msg: string) {
  toast.value = msg
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => { toast.value = '' }, 6000)
}
function onSave(blob: Blob, name: string) { downloadBlob(blob, name) }
function onError(err: Error) { notify(err.message || String(err)) }
async function onOpenRequest() {
  const files = await pickFile(accept.value)
  if (files[0]) await active.value?.open(files[0])
}
async function onSaveRequest() { await active.value?.saveAs() }
async function openLocal() { await onOpenRequest() }
function setLook(l: Look) { look.value = l }
</script>

<template>
  <div class="pg-root" :data-look="look" :data-layout="look === 'paper2016' ? 'ribbon' : 'bars'">
    <div class="pg-bar">
      <span class="pg-brand">VibeOffice <i>Vue</i></span>
      <button class="pg-tab" :class="{ on: tab === 'word' }" @click="tab = 'word'">Quire · Word</button>
      <button class="pg-tab" :class="{ on: tab === 'excel' }" @click="tab = 'excel'">Ledger · Spreadsheet</button>
      <button class="pg-tab" :class="{ on: tab === 'pptx' }" @click="tab = 'pptx'">Lectern · Presentation</button>
      <span class="pg-spacer"></span>
      <button class="pg-open" @click="openLocal">Open File…</button>
      <label class="pg-look">Look
        <select :value="look" @change="setLook(($event.target as HTMLSelectElement).value as Look)">
          <option value="classic">Classic</option>
          <option value="paper">Paper</option>
          <option value="paper2016">Paper 2016</option>
        </select>
      </label>
    </div>
    <div class="pg-body">
      <div v-show="tab === 'word'" class="pg-pane"><VWord ref="word" :look="look" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" /></div>
      <div v-show="tab === 'excel'" class="pg-pane"><VExcel ref="excel" :look="look" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" /></div>
      <div v-show="tab === 'pptx'" class="pg-pane"><VPptx ref="pptx" :look="look" @save="onSave" @error="onError" @open-request="onOpenRequest" @save-request="onSaveRequest" /></div>
    </div>
    <div v-if="toast" class="pg-toast">{{ toast }}</div>
  </div>
</template>

<style>
html, body, #app { height: 100%; margin: 0; }
.pg-root { height: 100%; display: flex; flex-direction: column; font: 12px/1.3 var(--ui-font, Tahoma, sans-serif); background: var(--dock-b, #C4DAFA); color: var(--ink, #000); }
.pg-bar { flex: none; display: flex; align-items: center; gap: 6px; padding: 5px 10px; background: linear-gradient(90deg, var(--dock-a, #9EBEF5), var(--dock-b, #C4DAFA)); border-bottom: 1px solid var(--chrome-line, #6f93cf); }
.pg-brand { font-weight: bold; margin-right: 10px; color: var(--brand-ink, #1f3f7f); }
.pg-brand i { font-style: normal; font-weight: normal; opacity: .7; }
.pg-tab { border: 1px solid transparent; background: none; padding: 4px 12px; border-radius: 6px 6px 0 0; cursor: pointer; font: inherit; color: inherit; }
.pg-tab:hover { background: var(--hot-a, #FFF4CC); }
.pg-tab.on { background: var(--surface, #fff); border-color: var(--chrome-line, #6f93cf); border-bottom-color: var(--surface, #fff); font-weight: bold; }
.pg-spacer { flex: 1; }
.pg-open { border: 1px solid var(--btn-border, #7C93BA); background: linear-gradient(var(--btn-a, #fff), var(--btn-b, #e7edf7)); padding: 3px 12px; border-radius: 4px; cursor: pointer; font: inherit; }
.pg-look { display: flex; align-items: center; gap: 4px; color: var(--muted, #4c5a72); }
.pg-look select { font: inherit; padding: 2px; }
.pg-body { flex: 1; min-height: 0; position: relative; }
.pg-pane { position: absolute; inset: 0; }
.pg-pane > * { height: 100%; }
.pg-toast { position: fixed; left: 50%; bottom: 40px; transform: translateX(-50%); background: var(--tip-bg, #ffffe1); border: 1px solid var(--tip-line, #000); padding: 8px 14px; box-shadow: 2px 2px 6px rgba(0,0,0,.3); z-index: 5000; max-width: 70vw; }
</style>
