# automation-report

[![Build Status](https://travis-ci.org/fczero/automation-report.svg?branch=master)](https://travis-ci.org/fczero/automation-report)
[![Coverage Status](https://coveralls.io/repos/github/fczero/automation-report/badge.svg?branch=master)](https://coveralls.io/github/fczero/automation-report?branch=master)

Jenkins Cucumber Report scraper to XLSX.
## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites
1. python 3.5+
1. python 3-venv
1. Must be connected to VPN
1. there should be no *myenv* folder on the script directory when running for the first time

### Usage

```sh
$python3 ccp_daily_automation.py [-h] [-s | -r | -t]
optional arguments:
         -h, --help        show this help message and exit
         -s, --smoke       Generate Smoke Report
         -r, --regression  Generate Regression Report
         -t, --test        Test mode
```

## Built With

* Python3
* OpenPyXL
* Requests

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Russell Delogu** - *Initial work*
* **Francis Lagadia** - *XLSX output, maintenance*

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## Acknowledgments

* Hat tip to anyone who's code was used
* Inspiration
* etc
