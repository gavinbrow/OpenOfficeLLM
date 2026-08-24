// Skills store: the catalog for the active host, plus create/delete.

import { create } from 'zustand'
import type { Skill } from '@openofficellm/shared'
import { getSkills, putSkill, deleteSkill, type SkillDraft } from '../api/client'
import { getHost } from '../host/bridge'

export interface SkillsState {
  skills: Skill[]
  userDir: string
  loading: boolean
  error: string | null
  load: () => Promise<void>
  save: (skill: SkillDraft) => Promise<boolean>
  remove: (id: string) => Promise<void>
}

export const useSkillsStore = create<SkillsState>((set, get) => ({
  skills: [],
  userDir: '',
  loading: false,
  error: null,

  load: async () => {
    set({ loading: true, error: null })
    try {
      // Filtered host-side by the detected host, so Excel never shows
      // "Rewrite formally" and Word never shows "Explain formula".
      const res = await getSkills(getHost())
      set({ skills: res.skills, userDir: res.userDir, loading: false })
    } catch (e) {
      set({ loading: false, error: (e as Error).message ?? 'Could not load skills.' })
    }
  },

  save: async (skill) => {
    try {
      await putSkill(skill)
      await get().load()
      return true
    } catch (e) {
      set({ error: (e as Error).message ?? 'Could not save the skill.' })
      return false
    }
  },

  remove: async (id) => {
    try {
      await deleteSkill(id)
      await get().load()
    } catch (e) {
      set({ error: (e as Error).message ?? 'Could not delete the skill.' })
    }
  },
}))
