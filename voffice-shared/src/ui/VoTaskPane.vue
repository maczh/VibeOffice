<script setup lang="ts">
import { icon } from './icons'
defineProps<{ title: string }>()
const emit = defineEmits<{ close: [] }>()
const closeSvg = icon('close', 10)
</script>

<template>
  <div class="taskpane">
    <div class="tp-head">
      <div class="tp-title"><span>{{ title }}</span></div>
      <button class="tp-nav" aria-label="Close" @click="emit('close')"><span v-html="closeSvg"></span></button>
    </div>
    <div class="tp-body"><slot></slot></div>
  </div>
</template>
