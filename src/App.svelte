<script lang="ts">
	import FileSelector from "./FileSelector.svelte";
	import * as XLSXLogic from "./XLSXLogic";
	

	export let name: string;
	var inputFile: FileList;
	var templateFile: FileList;

	var textField: HTMLParagraphElement;

	function onMatchBtnPressed() {
		switch (XLSXLogic.validateInput(inputFile, templateFile)) {
			case XLSXLogic.ErrorCode.MISSING_INPUT:
				textField.innerText = "missing input file!";
				break;
			case XLSXLogic.ErrorCode.MISSING_TEMPLATE:
				textField.innerText = "missing template file!";
				break;
			case XLSXLogic.ErrorCode.UNKNOWN:
				textField.innerText = "Unknown error!";
				break;
			case XLSXLogic.ErrorCode.NONE:
				textField.innerText = "";
				XLSXLogic.matchFiles(inputFile[0], templateFile[0])
				break;
		}
	}
</script>

<main>
	<h1>Hello <span class="name">{name}</span>!</h1>
	<p><span class="name">SAM</span>'s column matching.</p>

	<FileSelector bind:files={inputFile} />
	<FileSelector bind:files={templateFile} />
	<button on:click={onMatchBtnPressed}>MATCH!</button>

	<p bind:this={textField} />
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
