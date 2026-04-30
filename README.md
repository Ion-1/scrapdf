# ScraPDF
GUI for screenshotting URLs from your spreadsheet and compiling them into a PDF.

> [!NOTE]
> AI Usage: LLMs were used for brute-forcing some of the layout stuff and the usual autocomplete that PyCharm provides.

## Usage
The program currently works in its most basic form.
1. You are able to pick/edit URLs from a spreadsheet
2. Get them screenshotted with Playwright
3. Replace / Retry screenshots
4. Export them to a PDF

To run the program install `uv` and, in the project directory, run
```python
uv run python -O .\main.py
```

## Roadmap to 0.1.0
- [ ] Exporting run successes / failues
- [ ] Customize how playwright attempts retries, e.g. number of retries, ignoring http error codes, browser / extension choice, etc...
- [ ] Filters to run customizations on certain sites, e.g. on sites who always throw 404 codes whilst loading fine
- [ ] Async fetching
