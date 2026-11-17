<script setup lang="ts">
/* The Paper 2016 ribbon: .ribbon > .rb-bar (tabs) + .rb-panel (groups of commands). */
import { ref, computed } from 'vue'
import type { RibbonTab, ToolbarItem, ToolbarButton, MenuItem } from './types'
import { icon } from './icons'
import VoMenu from './VoMenu.vue'

const props = defineProps<{ tabs: RibbonTab[]; modelValue?: string }>()
const emit = defineEmits<{ 'update:modelValue': [string]; file: [] }>()

const on = ref(props.modelValue || props.tabs[0]?.id || '')
const collapsed = ref(false)
const peek = ref(false)
const menuItems = ref<MenuItem[]>([])
const menuPos = ref({ x: 0, y: 0 })
const menuOpen = ref(false)

const tab = computed(() => props.tabs.find((t) => t.id === on.value) || props.tabs[0])
function pick(id: string) {
  if (id === 'file') { emit('file'); return }
  on.value = id
  emit('update:modelValue', id)
  if (collapsed.value) peek.value = true
}
function run(b: ToolbarButton) { if (!b.disabled) b.run?.() }
function openMenu(b: ToolbarButton, e: MouseEvent) {
  if (!b.menu?.length || b.disabled) return
  const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
  menuItems.value = b.menu
  menuPos.value = { x: r.left, y: r.bottom }
  menuOpen.value = true
  e.stopPropagation()
}
const svg = (name?: string) => name ? icon(name, 24) : ''
const isBtn = (it: ToolbarItem): it is ToolbarButton => it.kind === 'button'
</script>

<template>
  <div class="ribbon" :class="{ collapsed, peek }">
    <div class="rb-bar">
      <div class="rb-tabs">
        <button class="rb-tab file" @click="pick('file')">File</button>
        <button v-for="t in tabs" :key="t.id" class="rb-tab" :class="{ on: on === t.id, ctx: t.ctx }"
          @click="pick(t.id)">{{ t.label }}</button>
      </div>
      <button class="rb-pin" title="Collapse the ribbon" @click="collapsed = !collapsed; peek = false">▴</button>
    </div>
    <div class="rb-panelwrap">
      <div v-if="tab" class="rb-panel" @mouseleave="peek = false">
        <div class="rb-groups">
          <div v-for="g in tab.groups" :key="g.label" class="rb-group">
            <div class="rb-body" :class="{ 'rb-col': g.rows }">
              <template v-for="(it, i) in g.items" :key="i">
                <span v-if="it.kind === 'sep'" class="tb-sep"></span>
                <span v-else-if="isBtn(it) && it.menu?.length" class="tb-split rb-large">
                  <button class="tb-btn" :class="{ on: it.on, dis: it.disabled }" :title="it.tip || it.label" @click="run(it)">
                    <span v-if="it.icon" v-html="svg(it.icon)"></span>
                    <span v-if="it.label" class="tb-lbl">{{ it.label }}</span>
                  </button>
                  <button class="tb-dd" @click="openMenu(it, $event)"><span class="dd-arrow"></span></button>
                </span>
                <button v-else-if="isBtn(it)" class="tb-btn rb-large" :class="{ on: it.on, dis: it.disabled, 'tb-drop': it.menu }"
                  :title="it.tip || it.label" @click="it.menu?.length ? openMenu(it, $event) : run(it)">
                  <span v-if="it.icon" v-html="svg(it.icon)"></span>
                  <span v-if="it.label" class="tb-lbl">{{ it.label }}</span>
                  <span v-if="it.menu?.length" class="dd-arrow"></span>
                </button>
              </template>
            </div>
            <div class="rb-label">{{ g.label }}
              <button v-if="g.launch" class="rb-launch" @click="g.launch()">⌄</button>
            </div>
          </div>
        </div>
      </div>
    </div>
    <VoMenu v-if="menuOpen" :items="menuItems" :x="menuPos.x" :y="menuPos.y" @close="menuOpen = false" />
  </div>
</template>
