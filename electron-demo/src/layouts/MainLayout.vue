<template>
  <q-layout view="lHh Lpr lFf">
    <q-header elevated>
      <q-toolbar>
        <q-toolbar-title> excel2json </q-toolbar-title>

        <div>v{{ curVersion }}</div>
      </q-toolbar>
    </q-header>

    <q-page-container>
      <router-view />
    </q-page-container>
  </q-layout>
</template>

<script lang="ts">
import { defineComponent, onMounted, ref } from 'vue';

export default defineComponent({
  name: 'MainLayout',

  setup() {
    const curVersion = ref('0.0.1');
    // onMounted(()=>{
    //   GetVersion();
    // });
    // function GetVersion(){

    // };
    return {
      curVersion,
    };
  },
});
</script>
