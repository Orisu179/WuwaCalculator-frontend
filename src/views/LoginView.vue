<script setup lang="ts">
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardContent, CardDescription, CardFooter, CardTitle, CardHeader } from '@/components/ui/card'
import { Label } from '@/components/ui/label';
import { ref } from 'vue'

const email = ref('');
const passwordInput = ref('');
</script>

<template>
  <Card class="w-full max-w-sm">
    <CardHeader>
      <CardTitle class="text-2xl">
        Login
      </CardTitle>
      <CardDescription>
        Enter your email below to login
      </CardDescription>
    </CardHeader>
    <CardContent class="grid gap-4">
      <div class="grid gap-2">
        <Label for="email"> Email </Label>
        <Input id="email" type="email" placeholder="your@email.com" required :value="email"/>
      </div>
      <div class="grid gap-2">
        <Label for="password"> Password </Label>
        <Input id="password" type="password" required :value="passwordInput"/>
      </div>
    </CardContent>
    <CardFooter>
      <Button class="w-full">
        Sign in
      </Button>
    </CardFooter>
  </Card>
  <Input type="email" placeholder="email"/>
</template>

<style scoped>

</style>