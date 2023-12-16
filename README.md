# Outlook Signature Generator

This is a lightweight script to prompt a user to setup a company signature for outlook following a template

It also has the capabilities to detect if a new template has been deployed and will then re prompt the user upon execution

## Templates

Its easiest to make the template within outlook using the subsitute key phrases:

These fields are populated automatically by looking up the user the script is running as within AD,
the user will have the ability to edit then within the gui

1. @FullName
1. @Title
1. @Address
1. @TNumber
    * *Telephone Number*
1. @MNumber
    * *Mobile Number*
1. @Flex
    * *Flexible work information*

Please see the examples folder for a rough example of how a signature may look

### Master Templates

These should be stored in a network location that all users have read access too or NETLOGON on the company Domain controllers as they need to be able to reach both for this script to work
