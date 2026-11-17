<script setup lang="ts">
/* Open from / save to an online place: WebDAV, S3 (SigV4), presigned URL, HTTP GET / POST.
 * The host app shows this dialog; on submit it gets a typed source/target for fetchOnline/saveOnline. */
import { ref, computed } from 'vue'
import VoDialog from './VoDialog.vue'
import type { OnlineSource, SaveTarget } from '../online/types'

const props = withDefaults(defineProps<{ mode: 'open' | 'save'; fileName?: string }>(), { mode: 'open' })
const emit = defineEmits<{ close: []; submit: [source: OnlineSource | SaveTarget] }>()

type Kind = 'webdav' | 's3' | 'presigned' | 'http' | 'post'
const kind = ref<Kind>('webdav')
const f = ref({
  url: '', username: '', password: '', token: '',
  endpoint: '', bucket: '', key: '', region: 'us-east-1', accessKeyId: '', secretAccessKey: '', sessionToken: '', forcePathStyle: true,
  putUrl: '', postUrl: '', field: 'file',
})

const kinds = computed(() => props.mode === 'open'
  ? [['webdav', 'WebDAV'], ['s3', 'S3'], ['presigned', 'Presigned URL'], ['http', 'HTTP URL']]
  : [['webdav', 'WebDAV'], ['s3', 'S3'], ['presigned', 'Presigned URL'], ['post', 'HTTP POST']])

const canSubmit = computed(() => {
  const v = f.value
  switch (kind.value) {
    case 'webdav': return !!v.url
    case 's3': return !!(v.endpoint && v.bucket && v.key && v.accessKeyId && v.secretAccessKey)
    case 'presigned': return !!v.url && (props.mode === 'open' || !!v.putUrl)
    case 'http': return !!v.url
    case 'post': return !!v.postUrl
  }
})

function submit() {
  const v = f.value
  let out: OnlineSource | SaveTarget
  switch (kind.value) {
    case 'webdav':
      out = { kind: 'webdav', url: v.url, username: v.username || undefined, password: v.password || undefined, token: v.token || undefined }
      break
    case 's3':
      out = { kind: 's3', endpoint: v.endpoint, bucket: v.bucket, key: v.key, region: v.region || undefined, accessKeyId: v.accessKeyId, secretAccessKey: v.secretAccessKey, sessionToken: v.sessionToken || undefined, forcePathStyle: v.forcePathStyle }
      break
    case 'presigned':
      out = props.mode === 'open' ? { kind: 'presigned', url: v.url } : { kind: 'presigned', url: v.url, putUrl: v.putUrl }
      break
    case 'http':
      out = { kind: 'http', url: v.url }
      break
    case 'post':
      out = { kind: 'post', url: v.postUrl, field: v.field || 'file' }
      break
  }
  emit('submit', out)
}
</script>

<template>
  <VoDialog :title="mode === 'open' ? 'Open from Online Location' : 'Save to Online Location'" :width="460" @close="emit('close')">
    <div class="tabs-head">
      <button v-for="[k, label] in kinds" :key="k" class="tab" :class="{ on: kind === k }" @click="kind = k as Kind">{{ label }}</button>
    </div>
    <div class="tabs-body">
      <template v-if="kind === 'webdav'">
        <div class="field wide"><label>URL</label><input v-model="f.url" type="text" placeholder="https://dav.example.com/files/report.docx" style="flex:1"></div>
        <div class="field wide"><label>Username</label><input v-model="f.username" type="text" style="flex:1"></div>
        <div class="field wide"><label>Password</label><input v-model="f.password" type="password" style="flex:1"></div>
        <div class="field wide"><label>Bearer token</label><input v-model="f.token" type="text" placeholder="(instead of username/password)" style="flex:1"></div>
      </template>
      <template v-else-if="kind === 's3'">
        <div class="field wide"><label>Endpoint</label><input v-model="f.endpoint" type="text" placeholder="https://s3.amazonaws.com" style="flex:1"></div>
        <div class="field wide"><label>Bucket</label><input v-model="f.bucket" type="text" style="flex:1"></div>
        <div class="field wide"><label>Object key</label><input v-model="f.key" type="text" :placeholder="fileName || 'documents/report.docx'" style="flex:1"></div>
        <div class="field wide"><label>Region</label><input v-model="f.region" type="text" style="flex:1"></div>
        <div class="field wide"><label>Access key ID</label><input v-model="f.accessKeyId" type="text" style="flex:1"></div>
        <div class="field wide"><label>Secret key</label><input v-model="f.secretAccessKey" type="password" style="flex:1"></div>
        <div class="field wide"><label>Session token</label><input v-model="f.sessionToken" type="text" placeholder="(optional)" style="flex:1"></div>
        <label class="ck"><input v-model="f.forcePathStyle" type="checkbox"> Path-style addressing (MinIO)</label>
      </template>
      <template v-else-if="kind === 'presigned'">
        <div class="field wide"><label>GET URL</label><input v-model="f.url" type="text" style="flex:1"></div>
        <div v-if="mode === 'save'" class="field wide"><label>PUT URL</label><input v-model="f.putUrl" type="text" style="flex:1"></div>
      </template>
      <template v-else-if="kind === 'http'">
        <div class="field wide"><label>URL</label><input v-model="f.url" type="text" placeholder="https://example.com/report.docx" style="flex:1"></div>
      </template>
      <template v-else>
        <div class="field wide"><label>POST URL</label><input v-model="f.postUrl" type="text" placeholder="https://example.com/upload" style="flex:1"></div>
        <div class="field wide"><label>Form field</label><input v-model="f.field" type="text" style="flex:1"></div>
      </template>
      <p class="tp-note">The server must allow this page's origin (CORS). Credentials stay in this page; nothing is sent anywhere else.</p>
    </div>
    <template #foot>
      <button class="btn primary" :disabled="!canSubmit" @click="submit">{{ mode === 'open' ? 'Open' : 'Save' }}</button>
      <button class="btn" @click="emit('close')">Cancel</button>
    </template>
  </VoDialog>
</template>
