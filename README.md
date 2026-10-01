# excel-data-cleaning-and-matching
Excel data cleaning, standardisation and multi-criteria matching using compound screening data.

# Excel Data Cleaning & Compound Matching

## Overview

This project demonstrates a small data-cleaning and data-matching workflow carried out in Excel using data from a compound screening experiment.

The starting point was a list of **14 screening hit compounds**, identified by their Well IDs. These hits needed to be matched back to a much larger compound library containing thousands of records, before reconnecting the identified compounds with supplier information.

The project demonstrates practical data-cleaning, standardisation, lookup and data-integration techniques in Excel.

## Dataset

The dataset contains compound identifiers, plate/well information and supplier identifiers.

The dataset has been anonymised for portfolio use. Compound, plate and supplier identifiers have been replaced with fictional identifiers while preserving the structure and relationships required to demonstrate the workflow.

No project-specific information or experimental conclusions are included.

## The Screening Hits

The 14 screening hits were identified by their Well IDs:

**A11, F4, G8, E8, G6, A3, D10, E3, C3, D9, C10, B9, G11 and C5.**

These Well IDs represented the positions of the hit compounds within the screening plates.

The challenge was to locate these hits within a much larger compound library and retrieve the corresponding compound information.

## Data Cleaning

The source data used a separate row identifier and column number to represent the position of each compound within a plate.

The original `COL` field contained well numbers with leading zeros, such as `02` and `03`.

When these values were combined directly with the row identifier, they produced Well IDs such as:

`A02`

However, the screening hit list used the standardised form:

`A2`

Changing the cell formatting did not solve the problem because the underlying values were stored as text. I therefore created a separate cleaned field using Excel's `VALUE()` function:

```excel
=VALUE(E5)
```

This converted:

`"02"` → `2`

The cleaned value was then combined with the row identifier using `CONCAT()`:

```excel
=CONCAT(D5,H5)
```

This produced standardised Well IDs such as `A2`, `A3`, `F4` and `G8`.

I retained the original column rather than modifying the source data directly, creating a separate derived field for the cleaned values.

## Matching the Compound Library

The standardised Well IDs were then matched against a compound library containing thousands of records.

A second identifier, the Barcode Plate, was also required because a Well ID such as `A2` could occur on more than one plate.

I therefore used `XLOOKUP()` with two matching conditions:

```excel
=XLOOKUP(1,
('Compound Library'!$G$5:$G$10244=E8)*
('Compound Library'!$F$5:$F$10244=$E$6),
'Compound Library'!$A$5:$Z$10244,
"Not found")
```

The two conditions checked:

1. Whether the Well ID matched the screening hit
2. Whether the Barcode Plate matched the relevant plate

Multiplying the two Boolean arrays produces `1` only where both conditions are true, allowing `XLOOKUP()` to identify the correct compound-library record.

## Reconnecting Supplier Data

The identified compounds were then matched with supplier information.

The supplier results were returned in a different order from the original compound library. A supplier compound identifier was therefore used as a common identifier to reconnect the supplier results to the original Well IDs.

This meant that the datasets did not need to be in the same order for the records to be correctly matched.

## Outcome

The workflow allowed me to:

* Start with 14 screening hit compounds
* Standardise inconsistent Well IDs
* Preserve the original source data
* Match the screening hits against a library containing thousands of records
* Use multiple criteria to identify the correct compound records
* Reconnect supplier results using a common compound identifier
* Avoid manually searching and rearranging thousands of records

## Skills Demonstrated

**Excel**

* Data cleaning
* Data standardisation
* `VALUE()`
* `CONCAT()`
* `XLOOKUP()`
* Boolean logic
* Multi-criteria lookups
* Data integration
* Working with datasets containing thousands of records

## What I Learned

This exercise highlighted the importance of understanding how Excel stores data, rather than relying only on how values appear on screen.

In particular, I learned that changing formatting does not necessarily change the underlying data type. In this case, `VALUE()` was used to convert text such as `"02"` into the numerical value `2`, allowing the resulting Well ID to match the existing `A2` format.

It also provided a practical introduction to using common identifiers to connect datasets that are stored in different orders.
