<script setup lang="ts">
/* One toolbar row: .toolbar > grip + buttons / split buttons / combos / separators. */
import { ref } from 'vue'
import type { ToolbarItem, ToolbarButton, ToolbarCombo, MenuItem } from './types'
import { icon } from './icons'
import VoMenu from './VoMenu.vue'

defineProps<{ items: ToolbarItem[]; grip?: boolean }>()

const menuFor = ref<string | null>(null)
const menuPos = ref({ x: 0, y: 0 })
const menuItems = ref<MenuItem[]>([])
const comboFocus = ref<string | null>(null)

function openMenu(b: ToolbarButton, e: MouseEvent) {
  if (!b.menu?.length || b.disabled) return
  const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
  menuItems.value = b.menu
  menuPos.value = { x: r.left, y: r.bottom }
  menuFor.value = b.id
  e.stopPropagation()
}
function run(b: ToolbarButton) { if (!b.disabled) b.run?.() }
function comboPick(c: ToolbarCombo, it: MenuItem) {
  menuFor.value = null
  if (it.label != null) c.onChange?.(it.label)
}
function comboOpen(c: ToolbarCombo, e: MouseEvent) {
  if (c.disabled || !c.items?.length) return
  const r = (e.currentTarget as HTMLElement).closest('.combo')!.getBoundingClientRect()
  menuItems.value = c.items
  menuPos.value = { x: r.left, y: r.bottom }
  menuFor.value = c.id
}
function comboKey(c: ToolbarCombo, e: KeyboardEvent) {
  if (e.key === 'Enter') { c.onChange?.((e.target as HTMLInputElement).value); (e.target as HTMLInputElement).blur() }
}
const svg = (name?: string) => name ? icon(name) : ''
</script>

<template>
  <div class="toolbar">
    <span v-if="grip !== false" class="tb-grip"></span>
    <template v-for="(it, i) in items" :key="i">
      <span v-if="it.kind === 'sep'" class="tb-sep"></span>
      <div v-else-if="it.kind === 'combo'" class="combo" :class="{ focus: comboFocus === it.id, dis: it.disabled }"
        :style="{ width: (it.width || 120) + 'px' }">
        <input class="combo-in" :value="it.value" :disabled="it.disabled" :readonly="!it.editable"
          @focus="comboFocus = it.id" @blur="comboFocus = null" @keydown="comboKey(it, $event)"
          @change="it.onChange?.(($event.target as HTMLInputElement).value)">
        <button class="combo-btn" tabindex="-1" @click="comboOpen(it, $event)"><span class="dd-arrow"></span></button>
      </div>
      <template v-else>
        <span v-if="it.menu?.length" class="tb-split">
          <button class="tb-btn" :class="{ on: it.on, dis: it.disabled, 'with-label': it.label }" :title="it.tip || it.label"
            @click="run(it)">
            <span v-if="it.icon" v-html="svg(it.icon)"></span>
            <span v-if="it.label" class="tb-lbl">{{ it.label }}</span>
          </button>
          <button class="tb-dd" tabindex="-1" @click="openMenu(it, $event)"><span class="dd-arrow"></span></button>
        </span>
        <button v-else class="tb-btn" :class="{ on: it.on, dis: it.disabled, 'with-label': it.label }" :title="it.tip || it.label"
          @click="run(it)">
          <span v-if="it.icon" v-html="svg(it.icon)"></span>
          <span v-if="it.label" class="tb-lbl">{{ it.label }}</span>
        </button>
      </template>
    </template>
    <VoMenu v-if="menuFor" :items="menuItems" :x="menuPos.x" :y="menuPos.y" @close="menuFor = null" />
  </div>
</template>
