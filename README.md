# automation-report
CCP Daily Automation Report

The report will be saved as an XLSX in the same directory.
1. Login to the VPN
 2.a     Add executable permission using chmod
         Ex. 
         
```bash
chmod 775 ccp_daily_automation.py
```

 3       Run using $./ccp_daily_automation.py
   OR
 2.b     Run using $python3 ccp_daily_automation.py
 ```
    Help
     usage: ccp_daily_automation.py [-h] [-s | -r | -t]

     Scrape Jenkins report and create XLSX report.

     optional arguments:
         -h, --help        show this help message and exit
         -s, --smoke       Generate Smoke Report
         -r, --regression  Generate Regression Report
         -t, --test        Test mode, Smoke Confirmation
```


# automation-report

Scrape Jenkins cucumber report and create XLSX report.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

What things you need to install the software and how to install them

```
python 3.4+
python 3-venv
```

### Installing

A step by step series of examples that tell you have to get a development env running

Say what the step will be

```
Give the example
```

And repeat

```
until finished
```

End with an example of getting some data out of the system or using it for a little demo

## Running the tests

Explain how to run the automated tests for this system

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Must be connected to VPN when run.

## Built With

* Python3
* OpenPyXl
* Requests


## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Russell Delogu** - *Initial work*
* **Francis Lagadia** - *XLSX output*

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## Acknowledgments

* Hat tip to anyone who's code was used
* Inspiration
* etc
