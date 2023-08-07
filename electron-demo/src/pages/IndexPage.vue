<template>
  <q-page class="row  justify-center">
  <div class="col-10 q-pa-md d-flex justify-center">
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Excel文件路径" @click="SelectExcelFile"/>
      <q-input  readonly filled v-model="excelPath" placeholder="选择Excel文件" :dense="dense" />
      
    </div>
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Json输出目录" @click="SelectOutputJsonDir"/>
      <q-input  readonly filled v-model="jsonPathDir" placeholder="选择Json输出目录" :dense="dense" />
      
    </div>
    <div class="d-flex justify-center align-center">
      <div class="text-center">
        <q-btn color="white" text-color="black" label="转换为json" @click="Translate2Json"/>
      </div>
    </div>
  </div>
</q-page>

</template>

<script lang="ts">
// import { transcode } from 'buffer';
// import { Todo, Meta } from 'components/models';
import { defineComponent,  ref } from 'vue';
// import { BrowserWindow } from 'electron';

export default defineComponent({
  name: 'IndexPage',
  
  setup () {
    var excelPath =ref('')
    var jsonPathDir=ref('')
   
    async function SelectExcelFile(){
     excelPath.value=await window.electronAPI.openExcelFile()
    };
    
    async function SelectOutputJsonDir(){
      jsonPathDir.value=await window.electronAPI.saveJsonDir()
    };
    async function Translate2Json(){
      await window.electronAPI.transform2json();
    };
    
    return {
       
       model: ref(null),
       text: ref(''),
       excelPath,
       jsonPathDir,
      dense: ref(false),
      SelectExcelFile,
      SelectOutputJsonDir,
      Translate2Json
      };
  }
});
</script>
