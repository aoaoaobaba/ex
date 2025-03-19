<template>
    <div>
        <button @click="callAjax">サーバー呼び出し</button>
    </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue';

export default defineComponent({
    methods: {
        callAjax() {
            fetch('/api/data')
                .then(response => response.json())
                .then(data => alert('データ: ' + JSON.stringify(data)))
                .catch(error => console.error('エラー:', error));
        }
    }
});
</script>