<template>
  <div id="app">
    <DownloadButton />
  </div>
</template>

<script>
import DownloadButton from './components/DownloadButton.vue'

export default {
  name: 'App',
  components: {
    DownloadButton
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
