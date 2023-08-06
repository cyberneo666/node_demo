<template>
  <q-page class="row  justify-center">
  <div class="col-10 q-pa-md d-flex justify-center">
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Excel文件路径" @click="SelectExcelFile"/>
      <q-input  readonly filled v-model="ph" placeholder="选择Excel文件" :dense="dense" />
      
    </div>
    <div class="q-mb-md">
      <q-btn color="white" text-color="black" label="选择Json输出目录" @click="SelectOutputJsonDir"/>
      <q-input  readonly filled v-model="ph" placeholder="选择Json输出目录" :dense="dense" />
      
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
import { defineComponent, onMounted, ref } from 'vue';
// import { BrowserWindow } from 'electron';

export default defineComponent({
  name: 'IndexPage',
  
  setup () {
    var ph =ref('')
    // onMounted(()=>{
    //   window.api.receive('selected-file',(args:string)=>{
    //     ph.value=args
    //      // ph.value='test'
    //     alert(ph.value)
    //   })
    // });
    async function SelectExcelFile(){
     // window.api.selectXlsxFile();
      ph.value=await window.electronAPI.openFile()
     

    };
    
    function SelectOutputJsonDir(){
      alert("hello!");
    };
    function Translate2Json(){
      alert("hello!");
    };
    
    return {
       
       model: ref(null),
       text: ref(''),
      ph,
      dense: ref(false),
      SelectExcelFile,
      SelectOutputJsonDir,
      Translate2Json
      };
  }
});
</script>
