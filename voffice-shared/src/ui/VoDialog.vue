<script setup lang="ts">
/* A dialog window: .dlg-overlay > .dlg with title bar, body and footer; draggable, Esc closes.
 * Modal by default; modeless lets the document keep working underneath. */
import { ref, onMounted, onBeforeUnmount } from 'vue'

const props = withDefaults(defineProps<{ title: string; width?: number; modeless?: boolean; primary?: string }>(), { width: 0, modeless: false })
const emit = defineEmits<{ close: []; ok: [] }>()

const dlg = ref<HTMLElement | null>(null)
const pos = ref<{ x: number; y: number } | null>(null)
let drag: { dx: number; dy: number } | null = null

function startDrag(e: MouseEvent) {
  const r = dlg.value!.getBoundingClientRect()
  drag = { dx: e.clientX - r.left, dy: e.clientY - r.top }
  pos.value = { x: r.left, y: r.top }
  window.addEventListener('mousemove', onDrag)
  window.addEventListener('mouseup', endDrag)
  e.preventDefault()
}
function onDrag(e: MouseEvent) {
  if (!drag) return
  pos.value = { x: Math.max(0, Math.min(innerWidth - 60, e.clientX - drag.dx)), y: Math.max(0, Math.min(innerHeight - 30, e.clientY - drag.dy)) }
}
function endDrag() { drag = null; window.removeEventListener('mousemove', onDrag); window.removeEventListener('mouseup', endDrag) }
function onKey(e: KeyboardEvent) { if (e.key === 'Escape') { emit('close'); e.stopPropagation() } }

onMounted(() => {
  document.addEventListener('keydown', onKey, true)
  const first = dlg.value?.querySelector<HTMLElement>('.dlg-body input, .dlg-body select, .dlg-body textarea, .dlg-body button')
  first?.focus()
})
onBeforeUnmount(() => { document.removeEventListener('keydown', onKey, true); endDrag() })
</script>

<template>
  <div class="dlg-overlay" :class="{ modeless }">
    <div ref="dlg" class="dlg" role="dialog" :aria-label="title"
      :style="{ width: width ? width + 'px' : undefined, left: pos ? pos.x + 'px' : '50%', top: pos ? pos.y + 'px' : '50%', transform: pos ? 'none' : 'translate(-50%, -50%)' }">
      <div class="dlg-title" @mousedown="startDrag">
        <span class="dlg-ttl">{{ title }}</span>
        <button class="dlg-x" aria-label="Close" @click="emit('close')">✕</button>
      </div>
      <div class="dlg-body"><slot></slot></div>
      <div v-if="$slots.foot" class="dlg-foot"><slot name="foot"></slot></div>
    </div>
  </div>
</template>
