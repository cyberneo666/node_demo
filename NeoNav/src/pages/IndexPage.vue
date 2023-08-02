<template>
  <q-page  style="height: 100vh;">
    <div class=" row justify-center items-center" 
    style="margin-top: 10%;height: 10vh; font-size: 300%;">
    
      {{ currentTime }}
  </div>
  <div class=" row justify-center " style="margin-top: 1%; height: 200px;min-width: 500px;">
    
    <q-input  style="min-width: 70vh;"
    outlined
    dense
    v-model="searchText"
    prefix-icon="arrow_drop_down"
    suffix-icon="search"
    @click:append="performSearch"
    @keyup.enter="performSearch"
    placeholder="Search"
  />
</div>
   
  </q-page>
  
  
</template>

<script lang="ts">
export default {
  data() {
    return {
      searchText: '',
      currentTime: ''
    };
  },
  mounted() {
    this.updateTime();
    setInterval(this.updateTime, 1000); // 每秒更新时间
  },
  methods: {
    updateTime() {
      const now = new Date();
      this.currentTime = now.toLocaleTimeString();
    },
    performSearch() {
      // 在这里执行搜索操作
      console.log('Performing search for:', this.searchText);
    }
  }
};

</script>
<style scoped>

</style>