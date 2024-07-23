import { createRouter, createWebHistory } from 'vue-router'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'home',
      component: () => import('../views/HomeView.vue')
    },
    {
      path: '/build',
      name: 'build',
      component: () => import('../views/BuildView.vue')
    },
    {
      path: '/composition',
      name: 'composition',
      component: () => import('../views/CompositionView.vue')
    },
    {
      path: '/weapon',
      name: 'weapon',
      component: () => import('../views/WeaponView.vue')
    },
    {
      path:'/echo',
      name: 'echo',
      component: () => import('../views/EchoView.vue')
    }
  ]
})

export default router
