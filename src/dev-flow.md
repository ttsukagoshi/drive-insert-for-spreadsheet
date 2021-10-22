# Dev Flow
1. Dev is to be conducted on branch `dev`
1. After being tested, the contents of `dev` will be merged into `main`, subsequently being versioned via `clasp` and deployed as the latest version of the editor add-on.
1. At the same time, `dev` should be merged into `boundScriptDev`, and after changing the const `IS_EDITOR_ADDON` to `false` and being tested as a spreadsheet-bound script, this `boundScriptDev` should, in turn, be merged into `main-script-bound`

Consequently, readers should look at either the `main` or `main-script-bound` branches for the latest version of the script. Note that the difference between the two branches is merely the boolean value of const `IS_EDITOR_ADDON`.