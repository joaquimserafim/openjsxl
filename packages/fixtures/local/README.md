# Local-only fixtures

Everything in this directory **except this README is git-ignored** (see the repo
`.gitignore`) and is never committed or published to npm. Use it to test the reader against
real `.xlsx` files we don't want in the committed corpus — typically because of their license.

Currently parked here (download them yourself to reproduce):

| File | Source | License | Why local-only |
| --- | --- | --- | --- |
| `SimpleWithComments.xlsx` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet/` | Apache-2.0 | legacy comments (`xl/comments1.xml` + VML); kept out of the MIT corpus |
| `comments.xlsx` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet/` | Apache-2.0 | same |

Tests that depend on these files must **skip gracefully when the file is absent**, so the
suite stays green on a fresh clone and in CI.
