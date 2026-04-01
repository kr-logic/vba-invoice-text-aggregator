# README

This repository contains an Excel VBA based data processing script made for business and office purposes.

The purpose of the code is to automatically aggregate and concatenate texts from invoice data, something the built-in Excel Pivot Table can't do.
Its goal is to extend the Pivot table with the additional aggregation, therefore you'll need a pivot table (or if already exists, a column that contains the invoiceCount of each invoice you are trying to aggregate).
The program assumes you have the invoice data on Work1 and its pivot on Work2, with column numbers being SPECIFIC to my use case, so please take this into consideration when refactoring and applying.

The program contains a few lines responsible for cleaning the text, this can be fully customized or removed.

The file includes both a basic version and an advanced version built to rely on sorting a little bit less. The advanced method takes longer to process, while the simple one is lightning fast.

The logic and coding are entirely my own development.

----------------------

Copyright (c) 2026 Princzinger Krisztián. All rights reserved.
