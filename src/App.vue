<script setup lang="ts">

import NavBar from '@/components/NavBar.vue'
</script>

<template>
  <div class="flex w-full flex-col min-h-screen">
    <NavBar />
    <RouterView />
  </div>
<!--      <nav>-->
<!--        <RouterLink to="/">Home</RouterLink>/-->
<!--        <RouterLink to="/about">About</RouterLink>-->
<!--      </nav>-->
</template>