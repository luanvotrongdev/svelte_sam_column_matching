<script lang="ts">
	import FileSelector from  './FileSelector.svelte'

	export let name: string;
</script>

<main>
	<h1>Hello <span class="name">{name}</span>!</h1>
	<p><span class="name">SAM</span>'s column matching.</p>

	<FileSelector></FileSelector>
	<FileSelector></FileSelector>
	<button>MATCH!</button>
</main>

<style>
	main {
		text-align: CENTER;
		padding: 1em;
		max-width: 240px;
		margin: 0 auto;
	}

	h1 {
		color: #000000;
		text-transform: uppercase;
		font-size: 4em;
		font-weight: 100;
	}

	.name {
		color: #ff3e00;
		text-transform: uppercase;
		font-size: 1em;
		font-weight: 100;
	}

	@media (min-width: 640px) {
		main {
			max-width: none;
		}
	}
</style>