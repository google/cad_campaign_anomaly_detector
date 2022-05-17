# CAD - Campaign Anomaly Detector

A read-only G-Ads script tool that identifies anomalies in key metrics of the accounts/campaigns, by comparing the daily/weekly average values of selected current and past periods, and alerts when current period metrics differ significantly. 


[User guide](https://docs.google.com/document/d/1PZZcCjLrg70d5Kj0Mr87ARAk2Q0dPook8950Fna6cSk/edit#heading=h.vgol1uz8ixf6)

</br>

# Project owners
- rsela@google.com
- omril@google.com
- meiravshaul@google.com
- Dev: 
- eladb@google.com

</br>

## Disclaimer

**This is not an officially supported Google product.**
Copyright 2021 Google LLC. This solution, including any related sample code or data, is made available on an “as is,” “as available,” and “with all faults” basis, solely for illustrative purposes, and without warranty or representation of any kind. This solution is experimental, unsupported and provided solely for your convenience. Your use of it is subject to your agreements with Google, as applicable, and may constitute a beta feature as defined under those agreements.  To the extent that you make any data available to Google in connection with your use of the solution, you represent and warrant that you have all necessary and appropriate rights, consents and permissions to permit Google to use and process that data.  By using any portion of this solution, you acknowledge, assume and accept all risks, known and unknown, associated with its usage, including with respect to your deployment of any portion of this solution in your systems, or usage in connection with your business, if at all.

</br>

# Key Components
The solutions includes the following components:
- Configuration Sheet (“input”) - Google Sheet to set the monitoring parameters (entities for monitoring, period selection, metrics & thresholds, email alerts)
- Simulator to support the period calculation and settings
- Dashboard Sheet (“results”) - dedicated tab in the configuration sheet, that presents the identified anomalies
- Data Studio Dashboard that presents the identified anomalies


## Video
[Link] (https://www.google.com/url?q=https://drive.google.com/file/d/1UyCuo_n9XQ6U9E1QSYaEBKKf9y6eYDgv/view?resourcekey%3D0-3I4BSEVo58qavsjVwF1pPQ&sa=D&source=docs&ust=1652738916206829&usg=AOvVaw1pQVmP4fBvjpDekksRGEBZ)


## Run

- Take this [spreadsheet] (https://docs.google.com/spreadsheets/d/17vN8ZRezhtQU11XKhOuexPkOBCLPiTO895oDJvhVblo/copy). 
- Copy its url for later.
- Rename it.
- Fill the thresholds in the “input” tab.
- Use the “simulator” tab to calculate the desired date ranges (optional).

[spreadsheet] (src/2022-05-17_00-12_1.png)

- Create a new script in G-Ads.
- Name it.
- Delete its template content. 
- Replace it with the code that appears in “Extensions >> Apps script”.
- Put the sheet’s url.

[spreadsheet] (src/2022-05-17_00-13.png)


- Run the script once (“preview”).
- Authorize the popup.
- Schedule the script to run every day.
- Results appear in the spreadsheet on the “Results” tab.


[spreadsheet] (src/2022-05-17_00-13_1.png)


</br>

## License
Apache Version 2.0
See [LICENSE](LICENSE)