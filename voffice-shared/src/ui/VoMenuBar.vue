<script setup lang="ts">
/* The Office-2003 menu bar: .menubar > .mb-item, one popup at a time, hover switches while open. */
import { ref } from 'vue'
import type { MenuDef } from './types'
import VoMenu from './VoMenu.vue'

const props = defineProps<{ menus: MenuDef[] }>()

const open = ref(-1)
const pos = ref({ x: 0, y: 0 })
const bar = ref<HTMLElement | null>(null)

function toggle(i: number, e: MouseEvent) {
  if (open.value === i) { open.value = -1; return }
  const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
  pos.value = { x: r.left, y: r.bottom }
  open.value = i
}
function hover(i: number) {
  if (open.value >= 0 && open.value !== i) {
    const row = bar.value?.children[i] as HTMLElement | undefined
    if (row) { const r = row.getBoundingClientRect(); pos.value = { x: r.left, y: r.bottom } }
    open.value = i
  }
}
</script>

<template>
  <div ref="bar" class="menubar" role="menubar">
    <div v-for="(m, i) in props.menus" :key="m.label" class="mb-item" :class="{ open: open === i }"
      role="menuitem" tabindex="0" @click="toggle(i, $event)" @mouseenter="hover(i)"
      @keydown.enter="toggle(i, $event as unknown as MouseEvent)">{{ m.label }}</div>
    <VoMenu v-if="open >= 0" :items="props.menus[open].items" :x="pos.x" :y="pos.y" @close="open = -1" />
  </div>
</template>
