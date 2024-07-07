<script setup lang="ts">
// export interface LinkProp {
//   title: string,
//   label?: string,
//   icon: string,
// }
// type NavProps = {
//   isCollapsed: boolean,
//   links: LinkProp[]
// }

// defineProps<NavProps>()
</script>

<template>
  <header class="sticky top-0 flex h-16 items-center gap-4 border-b bg-background px-4 md:px-6">
    <nav>
      <RouterLink to="/" class="text-muted-foreground transition-colors hover:text-foreground">
        Home
      </RouterLink>
    </nav>
  </header>
</template>

<style scoped>

</style>