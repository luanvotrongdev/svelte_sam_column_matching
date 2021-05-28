<script lang="ts">
import { bind } from "svelte/internal";
    var files : FileList;
</script>

<input type="file" id="file_input" bind:files>