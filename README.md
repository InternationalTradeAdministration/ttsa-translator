# NTTO TTSA Azure Function

This project provides an Azure Function that extracts Supply data from TTSA Excel Workbooks and puts it into a csv file.

## Prerequisites

Make sure you have the following installed.

* Python v3.8
* Pip
* Virtualenv
* Virtualenvwrapper
* Azure credentials

## Getting Started

  `git clone https://github.com/InternationalTradeAdministration/ttsa-translator.git`
  `cd ttsa-translator`
  `mkvirtualenv -r requirements.txt`
  `source /path/to/venv/bin/activate`

## Configuration

* You must have an Azure account and a storage account. 
* Create a function app at portal.azure.com
* Refer to deploy section for instructions on how to deploy to Azure
* Create a container in your storage account and change the path in function.json as needed

If using Visual Studio Code, you will need to install the following:
* Azure Account
* Azure App Service
* Azure CLI Tools
* Azure Functions
* Azure Storage

Make sure you have the following environment variables:
`AzureWebJobsStorage`, which should be the connection string to your storage account
`ContainerName`, which should be 'ntto'

## Run Locally

  `func host start`

You can upload Microsoft Excel files with the .xlsx extension to the container you specified in function.json. The function is triggered by new blobs that are uploaded into the specified container. After the conversion is done, a directory called translated will appear in the specified container and will hold the converted files.

## Deploy

* If you are using Visual Studio Code, navigate to the Azure tab on the Activity Bar on the left of the screen
* You should be able to see your active subscription in the functions section, and in your subscription you should see the function app you created earlier
* You should also see this function as a local project.
* Hover over the functions label and click "Deploy to Function App" button
* Follow the instructions at the top of VS Code and select the correct subscription and trigger when prompted.
* After the function deploys, run the command `func azure functionapp publish <your_functionapp_name>`
* The function will be published and you can upload files to the container you linked via portal.azure.com