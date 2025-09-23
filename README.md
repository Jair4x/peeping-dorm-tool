# Peeping Dorm Tool
A tool I made for one of my translation projects in Unity IL2CPP.

## The questions
### What (is this)? And what does it do?
As I said in the beginning, this is a tool I made for one of my translation projects.
[xUnity.AutoTranslator](https://github.com/bbepis/XUnity.AutoTranslator) isn't really my cup of tea for these kind of projects, way too overkill for what I want to do. And besides, I don't do MTL (Machine Translations), I love manual labor baby.

So, for my own convenience, I made this tool that gets the whole `Unity.Localization` bundle asset you can export with [UABEA](https://github.com/nesrak1/UABEA) by using "Export dump", and gets the `m_Id` and the `m_Localized` value and places them in a XLSX (Excel file) format for your convenience when translating.

Then, we "rebuild" the files. (We just make the files a .json with the ID and localized string that we can use with a mod to inject the lines directly in the game using these modified files)
### Where?
On your PC, obviously.
### When?
When what? When is the translation coming out? Lost all of it by accident when adapting [the mod](https://github.com/Jair4x/dorm-manager-patch) from MelonLoader to BepInEx for easier exportation.
Few months after this repo I'll probably be finishing the translation.

## Install
Install the [`ExcelJS`](https://www.npmjs.com/package/exceljs) library. Done. Everything else is from the base kit.
## Usage
Simple, just use your js/ts interpreter of preferece (node, bun run, deno run, whatever) and run the command.
Giving no args will default to these folders:

- Raw .json folder: "Raws"
- Extracted Excel files: "Extracted"
- Rebuilt .json files: "Rebuilt"
### Extract 
`{interpreter} ./extract.ts [-i "RawFolder"] [-o "ExcelFolder"]`
### Rebuild
`{interpreter} ./rebuild.ts [-i "ExcelFolder"] [-o "RebuiltFolder"]`
