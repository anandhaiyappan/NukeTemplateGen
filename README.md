# NukeTemplateGen
This App can be used to create a base nuke script quickly for shots using master shot's nuke script as a template.

## Screenshot

![NukeTemplateGen](https://user-images.githubusercontent.com/51224384/192089551-b35285c2-6444-44b8-abdb-ddc602beb966.jpg)

## Installation

* Clone the [NukeTemplateGen](https://github.com/anandhaiyappan/NukeTemplateGen.git) git repo in any location

```
https://github.com/anandhaiyappan/NukeTemplateGen.git
```

* Navigate to the folder and run the `main.py` using python.exe


## How to use

* Click on the `Sample Excel` button to create sample excel.
* Fill the excel with Shotname, First and Last frames
* Click on the `Sample Nuke File`  to create sample nuke script.
* Open the nuke script and edit project settings and add nuke nodes which you want for initial setup for the artist (ex:plates, grain, bg, Writenode template) and save the nuke file.
* Open the UI and click the `Load nk script` then fill the Shotname. First and Last frames will be fetched from the nuke script.
* Clieck the `Load Excel` it will process and fill number of shots available in the excel, you can choose between all the shots or number of shots then type the number based on your requirements
* Click the `OutPath` and selct the location for the all the nuke scripts to be created.
* Click the `Create Script` for creating the scripts.
* Click the `Open Outpath` to open the folder.

