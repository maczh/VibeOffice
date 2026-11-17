<script setup lang="ts">
defineProps<{ label?: string }>()
</script>

<template>
  <div class="busy"><div class="busy-box"><span class="busy-spin"></span><span>{{ label || 'Working…' }}</span></div></div>
</template>
