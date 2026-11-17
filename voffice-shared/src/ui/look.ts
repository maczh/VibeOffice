/* The look setting ('vibeoffice.look' in localStorage): classic, paper, paper2016.
 * The component root carries data-look / data-layout, as the original pages put them on <html>. */
import { ref, watch, type Ref } from 'vue'

export type Look = 'classic' | 'paper' | 'paper2016'
const KEY = 'vibeoffice.look'

export function getLook(): Look {
  try {
    const v = localStorage.getItem(KEY)
    if (v === 'paper' || v === 'paper2016') return v
  } catch { /* private mode */ }
  return 'classic'
}

export function setLook(l: Look): void {
  try { localStorage.setItem(KEY, l) } catch { /* private mode */ }
}

/** a shared reactive look; bind data-look/data-layout on the component root to it */
export function useLook(initial?: Look): { look: Ref<Look>; layout: Ref<string> } {
  const look = ref<Look>(initial || getLook())
  const layout = ref(look.value === 'paper2016' ? 'ribbon' : 'bars')
  watch(look, (l) => { layout.value = l === 'paper2016' ? 'ribbon' : 'bars'; setLook(l) })
  return { look, layout }
}
