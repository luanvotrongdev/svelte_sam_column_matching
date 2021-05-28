<script lang="ts">
import { bind } from "svelte/internal";
    export var files : FileList = null;
</script>

<input type="file" id="file_input" bind:files>