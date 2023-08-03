<template>
  <q-page class="row  justify-center">
  <div class="col-10 q-pa-md d-flex justify-center">
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Excel文件路径" @click="SelectExcelFile"/>
      <q-input filled v-model="ph" placeholder="选择Excel文件" :dense="dense" />
      
    </div>
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Json输出目录" @click="SelectOutputJsonDir"/>
      <q-input filled v-model="ph" placeholder="选择Json输出目录" :dense="dense" />
      
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
import { transcode } from 'buffer';
import { Todo, Meta } from 'components/models';
import { defineComponent, onMounted, ref } from 'vue';

export default defineComponent({
  name: 'IndexPage',
  
  setup () {
    
    function SelectExcelFile(){
      window.electron.ipcRenderer.send('open-file-dialog-for-xlsx')
    };
    function SelectOutputJsonDir(){
      alert("hello!");
    };
    function Translate2Json(){
      alert("hello!");
    };
    onMounted(()=>{
      window.electron.ipcRenderer.on('selected-file', (event, path) => {
      console.log('Selected file', path);
    // Do something with the selected file path...
  });
    });
    return {
       
       model: ref(null),
       text: ref(''),
      ph: ref(''),
      dense: ref(false),
      SelectExcelFile,
      SelectOutputJsonDir,
      Translate2Json
      };
  }
});
</script>
