I have a Python Playwright scraper file named `Scrapping_live_v1.py` and an Excel file named `gair malik.xlsx`.

I need you to modify the scraper and return the COMPLETE updated Python file.

## Goal

The scraper already extracts land details from mutation applications:

* Halka
* Mauja
* Thana
* Khata No
* Plot No
* Area1
* Area2

The extracted values appear in logs like:

```python
land details: (
    'BHAISHWAN',
    'जगपुरा',
    'Patna',
    '174',
    '225',
    '0',
    '2.656'
)
```

## New Requirement

Load `gair malik.xlsx` once when the script starts.

Add a permanent lookup system.

For every mutation application scraped, compare the extracted land details against records in `gair malik.xlsx`.

Add a new output column:

```python
GAIR_MALIK
```

Return:

```python
YES
```

if the land exists in the Gair Malik file.

Otherwise:

```python
NO
```

---

## Structure of gair malik.xlsx

The Excel contains rows like:

| Mauja / Thana No | Khata | Khesra | Rakba |
| ---------------- | ----- | ------ | ----- |
| DAULATPUR 221    | 72    | 146    | 0.01  |
| DAULATPUR 221    | 72    | 147    | 0.15  |
| DAULATPUR 221    | 72    | 152    | 0.05  |

Important:

Column 1 contains:

```text
Mauja + Thana Number
```

combined into a single field.

Column 2 contains:

```text
Khata
```

Column 3 contains:

```text
Khesra (Plot Number)
```

Column 4 contains Rakba and is not needed for matching.

---

## Matching Logic

Do NOT use Rakba.

Use:

```python
Mauja
Khata
Plot No
```

for matching.

Reason:

The scraper's Thana field contains values such as:

```text
Patna
```

while the Excel stores:

```text
221
```

so Thana values are not compatible.

Therefore match using:

```python
Normalized Mauja
Khata No
Plot No
```

only.

---

## Mauja Normalization

The portal may return:

```text
भैंसवाँ
```

while the Excel may contain:

```text
BHAISHWAN
```

Create a normalization system.

Requirements:

* trim spaces
* uppercase English text
* remove extra whitespace
* support Hindi and English variants
* allow configurable alias mapping

Example:

```python
MAUJA_ALIASES = {
    "भैंसवाँ": "BHAISHWAN",
    "BHAISHWAN": "BHAISHWAN",
    "BHAISHWAN": "BHAISHWAN"
}
```

Apply aliases before matching.

---

## Multiple Plot Numbers

The scraper may produce:

```text
146;147;152
```

for Plot No.

Split by:

```python
;
```

and if ANY plot number matches a Gair Malik record:

```python
GAIR_MALIK = YES
```

Otherwise:

```python
GAIR_MALIK = NO
```

---

## Required Functions

Create:

```python
load_gair_malik()
normalize_mauja()
is_gair_malik()
```

Load the Excel only once at startup.

Store records in a fast lookup structure such as:

```python
set()
```

for O(1) lookups.

---

## Output Changes

Add a new Excel column:

```python
GAIR_MALIK
```

Append it to every exported row.

Example output:

| Mauja     | Khata | Plot | GAIR_MALIK |
| --------- | ----- | ---- | ---------- |
| DAULATPUR | 72    | 146  | YES        |
| DAULATPUR | 72    | 999  | NO         |

---

## Debug Logging

Add:

```python
print(
    f"[GAIR MALIK CHECK] "
    f"Mauja={mauja} "
    f"Khata={khata} "
    f"Plot={plot}"
)
```

and

```python
print("[GAIR MALIK] MATCH FOUND")
```

or

```python
print("[GAIR MALIK] NO MATCH")
```

for debugging.

---

## Deliverable

Return the FULL modified `Scrapping_live_v1.py` file with all changes integrated.

Do not give patch snippets.

Do not give partial code.

Return the complete updated file.
