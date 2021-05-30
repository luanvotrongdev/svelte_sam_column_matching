<script lang="ts">
    export var files : FileList = null;
</script>

<input type="file" id="file_input" accept=".xlsx,.xls" bind:files>


<!-- <button on:click={document.getElementById('file-input').click}>Open</button>
<input id="file-input" type="file" name="name" style="display: none;" /> -->