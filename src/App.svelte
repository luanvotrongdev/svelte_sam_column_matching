<script lang="ts">
	import FileSelector from  './FileSelector.svelte'

	export let name: string;
	var inputFile : FileList;
	var templateFile : FileList;

	var textField : HTMLParagraphElement;

	function onMatchBtnPressed()
	{
		if(inputFile === null)
		{
			textField.innerText = "input file is missing"
			return 
		}
		if(templateFile === null)
		{
			textField.innerText = "template file is missing"
			return 
		}
			textField.innerText = ""

		// console.log(inputFile[0].name);
		// console.log(templateFile[0].name);
		
	}
</script>

<main>
	<h1>Hello <span class="name">{name}</span>!</h1>
	<p><span class="name">SAM</span>'s column matching.</p>

	<FileSelector bind:files={inputFile}></FileSelector>
	<FileSelector bind:files={templateFile}></FileSelector>
	<button on:click={onMatchBtnPressed}>MATCH!</button>

	<p bind:this={textField}></p>
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