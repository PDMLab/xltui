# xltui - minimal Terminal UI for Excel files

Minimal Excel Terminal UI (TUI) inspired by TUIs like

* [doxx](https://github.com/bgreenwell/doxx)
* [lazysql](https://github.com/jorgerojas26/lazysql)

## Status

This is the very beginning of this little project. Right now it's more of a CLI to explore some ideas and concepts.

What you can do already:

Open a `.xslx` file and render a particular sheet as a table.

![](assets/screenshot-xltui-helloworld.png)

Open a `.xlsx` file and render it as a tree.

![](assets/screenshot-xltui-tree.png)

More options:

```bash
xltui render --file sample.xlsx
xltui render --file sample.xlsx --sheet People --style table
xltui render --file sample.xlsx --sheet People --style tree --group-by Department
xltui render --file sample.xlsx --sheet-index 1 --columns Name,Email,Dept
xltui render --file sample.xlsx --json
xltui render --file sample.xslx --json --columns Name,Email
```

The `--json` option can be useful if you want to process the data e.g. using `jq`:

```bash
# this will export a single column "Amount" and jq will transform the array to an array of strings:
XlTui render --file HelloWorld.xlsx --json --columns Amount | jq -c '[ .[] | .[] | to_entries[] | (.value // "") | tostring ]' 
```




