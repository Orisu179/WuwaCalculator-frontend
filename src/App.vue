<script setup lang="ts">
import { RouterLink, RouterView } from 'vue-router'
import { useColorMode } from '@vueuse/core'
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from '@/components/ui/dropdown-menu'
import { Icon } from '@iconify/vue'

const mode = useColorMode();
</script>

<template>
  <div class="flex items-center justify-between">
    <DropdownMenu>
      <DropdownMenuTrigger as-child>
        <Button variant="outline">
          <Icon icon="radix-icons:moon" class="h-[1.2rem] w-[1.2rem] rotate-0 scale-100 transition-all dark:-rotate-90 dark:scale-0" />
          <Icon icon="radix-icons:sun" class="absolute h-[1.2rem] w-[1.2rem] rotate-90 scale-0 transition-all dark:rotate-0 dark:scale-100" />
          <span class="sr-only"> Toggle Theme </span>
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end">
        <DropdownMenuItem @click="mode = 'light'">
          Light
        </DropdownMenuItem>
        <DropdownMenuItem @click="mode = 'dark'">
          Dark
        </DropdownMenuItem>
        <DropdownMenu @click="mode = 'auto'">
          System
        </DropdownMenu>
      </DropdownMenuContent>
    </DropdownMenu>
  </div>
<!--      <nav>-->
<!--        <RouterLink to="/">Home</RouterLink>/-->
<!--        <RouterLink to="/about">About</RouterLink>-->
<!--      </nav>-->

  <RouterView />
</template>