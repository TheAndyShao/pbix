# pbix
A package providing tools to Power BI developers working with thin reports.

## Description
Most enterprise Power BI environments will choose to organise assets into a central data model, with connected "thin" reports.

Whilst separation of the reports from the data model brings many benefits, it also creates additional challenges for developers when trying to propagate changes from the data model to the associated thin reports.

This packages aims to provide tools to Power BI developers to overcome these challenges.

## Installation
Use package manager pip to install pbix from this repository

```bash
pip install git+https://github.com/TheAndyShao/pbix
```

## Usage
Refer to the included codes samples to see how common functionality can be utilised.


## Limitations
- Page level filter field replacement is currently not supported
- Customised fields like display names are not affected

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.