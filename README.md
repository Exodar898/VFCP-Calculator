# VFCP-Calculator

This is an Excel VBA Project which was created to complete calculations for the IRS Voluntary Fiduciary Correction Program.
Retirement plan Fiduciaries routinely fail to invest clients' funds into the correct sources by the required dates.
The IRS created the VFCP to allow fiduciaries to independently assess the damages owed to clients for lost earnings on these funds.
The IRS has provided a tool on their website to calculate the damages, but the tool requires that each account be enterred manually, one at a time.

The firm I worked for required a settlement to be calculated in which several thousand participants were involved. The manual data entry would've taken far too long.

Thus, I was asked to create this project which perfectly follows the rules of the VFCP but also allows an indeterminant amount of participants to be enterred at once.

For your convenience, I have extracted the VBA code and included it as a separate file, just in case you didn't feel like opening a macro enabled excel project from somebody who you've never met.

In addition, I have included source files for all data pulled as part of the project. Most of the data had to be pulled through HTML extractions of the dol.gov website listed below. This meant that I had to spend a fair amount of time grooming the data and creating a manual index to reference the data. Luckily, the project is structured such that only three or four lines of data need to be added each year, simply at the top of the table in the VFCP-Table tab. These new pieces of data are released by the DOL.

For more information on the Voluntary Fiduciary Correction Program, see https://www.dol.gov/agencies/ebsa/employers-and-advisers/plan-administration-and-compliance/correction-programs/vfcp
