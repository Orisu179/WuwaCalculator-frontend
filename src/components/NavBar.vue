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
type navProps = {
  name: string,
  path: string
}

// defineProps<NavProps>()
import ToggeThemeButton from '@/components/ToggeThemeButton.vue'
import {
  NavigationMenu,
  NavigationMenuItem,
  NavigationMenuLink,
  navigationMenuTriggerStyle
} from '@/components/ui/navigation-menu'
import { MenubarSeparator } from 'radix-vue'

const menuPaths : navProps[] = [
  {
    name: "Home",
    path: "/"
  },
  {
    name: "Build",
    path: "/build"
  },
  {
    name: "Composition",
    path: "/composition"
  },
  {
    name: "Echo",
    path: "/echo"
  },
  {
    name: "Weapon",
    path: "/weapon"
  }
]
</script>

<template>
  <header class="sticky top-0 flex justify-between h-16 items-center gap-4 border-b bg-background px-4 md:px-6">
    <div class="mx-2 flex flex-row">
      <NavigationMenu v-for="(menuItem, index) in menuPaths" :key="index">
<!--        <NavigationMenuItem>-->
          <RouterLink :to="menuItem.path">
            <NavigationMenuLink :class="navigationMenuTriggerStyle()">
              {{menuItem.name}}
            </NavigationMenuLink>
          </RouterLink>
<!--        </NavigationMenuItem>-->
        <MenubarSeparator />
      </NavigationMenu>
    </div>
    <ToggeThemeButton />
  </header>
</template>

<style scoped>

</style>