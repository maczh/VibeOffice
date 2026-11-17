<script setup lang="ts">
import { onMounted, onBeforeUnmount } from 'vue'
const props = withDefaults(defineProps<{ message: string; ms?: number }>(), { ms: 4000 })
const emit = defineEmits<{ done: [] }>()
let timer: ReturnType<typeof setTimeout>
onMounted(() => { timer = setTimeout(() => emit('done'), props.ms) })
onBeforeUnmount(() => clearTimeout(timer))
</script>

<template>
  <div class="toast" role="status">{{ message }}<slot></slot></div>
</template>
