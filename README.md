# The Azure DevOps Onboarding Api

The Azure DevOps Onboarding Api is an abstraction layer on top of the Azure DevOps API and is written in typescript. It is offered as a node package so that you can use it in your own jaascript/typescript script or Azure DevOps extension to ease Azure DevOps onboarding. It is part of three git repositories. See [Git Repository](https://github.com/JoostVoskuil/azure-devops-onboarding/README.md) for more information.

## Contribute

I would love to see pull-requests on this node package. Please contact me if you want to contribute.

## Now to use

Run: 'npm install azure-devops-onboarding --save'

Initialize onboaring:
'const onboarding: OnboardingServices = new OnboardingServices(AZUREDEVOPS_PAT, AZUREDEVOPS_PAT_OWNER, MS_GRAPH_APP_SECERET);'

You can use methods for example:
'await onboarding.group().createGroup(project, SecurityGroupName, description)' to create a group.

The constructor requires the secret settings like the Azure DevOps PAT, the PAT owner and the MS_GRAPH App secret. See [Git Repository](https://github.com/JoostVoskuil/azure-devops-onboarding/README.md) for more information about how to connect to Azure.

## Access Configuration

You can access the configuration by using 'onboarding.configuration().'

## Access Services

You can access services by using 'onboarding.<service()'

## Design considerations

It was a choice not to build the most efficient API but it was the choice to make it easy to use. It works with names (strings) instead of objects or id's. In this way it is user friendly but it hits the API harder because it constantly looks up information.

## Throttling

When using the Azure DevOps api you can get throttled because you are exceeding the rate limit. [See the documentation of Microsoft about this subject](https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/rate-limits?view=azure-devops)