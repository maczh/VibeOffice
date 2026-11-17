<script setup lang="ts">
/* A popup menu: the .menu / .m-item structure of the original chrome (common/ui.css). */
import { ref, computed, onMounted, onBeforeUnmount, nextTick } from 'vue'
import type { MenuItem } from './types'
import { icon } from './icons'

const props = defineProps<{ items: MenuItem[]; x: number; y: number; overDlg?: boolean }>()
const emit = defineEmits<{ close: [] }>()

const el = ref<HTMLElement | null>(null)
const active = ref(-1)
const subFor = ref(-1)
const subPos = ref({ x: 0, y: 0 })
const pos = ref({ x: props.x, y: props.y })

const enabled = (it: MenuItem) => !it.disabled && !it.sep && !it.head

function pick(i: number) {
  const it = props.items[i]
  if (!it || !enabled(it)) return
  if (it.sub?.length) { openSub(i); return }
  emit('close')
  it.run?.()
}

function openSub(i: number) {
  const it = props.items[i]
  if (!it?.sub?.length || it.disabled) { subFor.value = -1; return }
  const row = el.value?.children[i] as HTMLElement | undefined
  if (!row) return
  const r = row.getBoundingClientRect()
  subFor.value = i
  subPos.value = { x: r.right - 2, y: r.top - 3 }
}

function hover(i: number) {
  active.value = i
  const it = props.items[i]
  if (it?.sub?.length && !it.disabled) openSub(i)
  else subFor.value = -1
}

function onKey(e: KeyboardEvent) {
  const n = props.items.length
  if (e.key === 'Escape') { emit('close'); e.stopPropagation() }
  else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    const dir = e.key === 'ArrowDown' ? 1 : -1
    let i = active.value
    for (let k = 0; k < n; k++) { i = (i + dir + n) % n; if (enabled(props.items[i])) break }
    active.value = i
    subFor.value = -1
    e.preventDefault()
  } else if (e.key === 'Enter' || e.key === ' ') { pick(active.value); e.preventDefault() }
  else if (e.key === 'ArrowRight') { if (props.items[active.value]?.sub?.length) openSub(active.value); e.preventDefault() }
  else if (e.key === 'ArrowLeft') { emit('close'); e.preventDefault() }
}

function onDocDown(e: MouseEvent) {
  if (el.value && !el.value.contains(e.target as Node)) emit('close')
}

onMounted(async () => {
  await nextTick()
  const r = el.value!.getBoundingClientRect()
  let x = props.x, y = props.y
  if (x + r.width > innerWidth - 4) x = Math.max(4, innerWidth - 4 - r.width)
  if (y + r.height > innerHeight - 4) y = Math.max(4, props.y - r.height)
  pos.value = { x, y }
  el.value!.focus()
  document.addEventListener('mousedown', onDocDown, true)
})
onBeforeUnmount(() => document.removeEventListener('mousedown', onDocDown, true))

const mark = (it: MenuItem) => it.checked ? (it.radio ? '●' : '✓') : ''
const iconSvg = computed(() => (name?: string) => name ? icon(name) : '')
</script>

<template>
  <div ref="el" class="menu" :class="{ 'over-dlg': overDlg }" tabindex="-1" role="menu"
    :style="{ left: pos.x + 'px', top: pos.y + 'px' }" @keydown="onKey">
    <template v-for="(it, i) in items" :key="i">
      <div v-if="it.sep" class="m-sep"></div>
      <div v-else-if="it.head" class="m-head">{{ it.head }}</div>
      <div v-else class="m-item" :class="{ act: active === i, dis: it.disabled, bold: it.bold }" role="menuitem"
        @mouseenter="hover(i)" @click.stop="pick(i)">
        <span class="m-ico" :class="{ on: it.checked && !it.icon }">
          <span v-if="it.icon" v-html="iconSvg(it.icon)"></span>
          <span v-else-if="it.checked" class="m-check" :class="{ 'm-radio': it.radio }">{{ mark(it) }}</span>
        </span>
        <span class="m-lbl">{{ it.label }}</span>
        <span class="m-key">{{ it.key || '' }}</span>
        <span class="m-sub">{{ it.sub?.length ? '▶' : '' }}</span>
      </div>
    </template>
    <VoMenu v-if="subFor >= 0 && items[subFor]?.sub" :items="items[subFor].sub!" :x="subPos.x" :y="subPos.y"
      :over-dlg="overDlg" @close="subFor = -1" />
  </div>
</template>
