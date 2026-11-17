<script setup lang="ts">
import type { StatusCell } from './types'
defineProps<{ cells: StatusCell[] }>()
</script>

<template>
  <div class="statusbar">
    <div v-for="(c, i) in cells" :key="i" class="sb-cell"
      :class="{ grow: c.grow, flag: c.flag, on: c.on, act: c.act }" :title="c.tip"
      @click="c.run?.()">{{ c.text }}</div>
    <slot></slot>
  </div>
</template>
