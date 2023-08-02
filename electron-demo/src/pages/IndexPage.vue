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
import { defineComponent, ref } from 'vue';

export default defineComponent({
  name: 'IndexPage',
  
  setup () {
    const todos = ref<Todo[]>([
      {
        id: 1,
        content: 'ct1'
      },
      {
        id: 2,
        content: 'ct2'
      },
      {
        id: 3,
        content: 'ct3'
      },
      {
        id: 4,
        content: 'ct4'
      },
      {
        id: 5,
        content: 'ct5'
      }
    ]);
    const meta = ref<Meta>({
      totalCount: 1200
    });
    function SelectExcelFile(){
      alert("hello!");
    };
    function SelectOutputJsonDir(){
      alert("hello!");
    };
    function Translate2Json(){
      alert("hello!");
    };
    return {
       todos, meta ,
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
