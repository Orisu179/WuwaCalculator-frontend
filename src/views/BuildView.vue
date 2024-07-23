<script setup lang="ts">
import { ResizablePanelGroup, ResizableHandle, ResizablePanel} from '@/components/ui/resizable'
</script>

<template>
    <ResizablePanelGroup
      id="separator"
      direction="horizontal"
      class="w-screen min-h-full border rounded-md h-96"
    >
      <ResizablePanel id="left" :default-size="300" class="max-h-full md:max-h-screen">
        <div class="flex h-96 items-center justify-center">
          <h1>
            This is where the enter section is
          </h1>
        </div>
      </ResizablePanel>
      <ResizableHandle with-handle />
      <ResizablePanel id="right" :default-size="1000" class="max-h-full md:max-h-screen">
        <div class="flex h-96 items-center justify-center">
          <h1>
            This is where the editable section is
          </h1>
        </div>
      </ResizablePanel>
    </ResizablePanelGroup>
</template>

<style scoped>

</style>