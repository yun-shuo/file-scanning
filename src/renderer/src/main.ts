import './assets/main.css'

import { createApp } from 'vue'
import App from './App.vue'
import PrimeVue from 'primevue/config'
import Aura from '@primevue/themes/aura'
import lara from '@primevue/themes/lara'
import nora from '@primevue/themes/nora'
import material from '@primevue/themes/material'
import 'primeicons/primeicons.css'
import 'primeflex/primeflex.css'
// import 'primeflex/themes/primeone-light.css' // 亮系主题
import 'primeflex/themes/primeone-dark.css' // 暗系主题

import ToastService from 'primevue/toastservice'
import 'uno.css'
createApp(App)
  .use(PrimeVue, {
    theme: {
      preset: lara
    }
  })
  .use(ToastService)
  .mount('#app')
