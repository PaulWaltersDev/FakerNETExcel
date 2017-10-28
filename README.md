
# FakerNETExcel
An implementation of functions for the creation of fake customer test data in Faker.NET (https://www.nuget.org/packages/Faker.Net) to Excel 2013 (as an XLL add-in) and later. Useful when customer production data is banned or inaccessible and the manual creation of customer data is too laborious.

<h2>Installation</h2>

1) Download the package zip. Open and find the files -

FakerNETExcel-AddIn-packed.xll

FakerNETExcel-AddIn64-packed.xll

2) Open Excel 2013 or later. Create a new sheet.

3) Select File -> Options. In Excel Options, select "Add-Ins"

4) Next to the dropdown option "Manage", selected option "Excel Add-Ins", click "Go..."

5) In the Add-Ins popup, click Browse and open the XLL addin file above corresponding to your Excel installation.

6) Click the checkbox for the added file and then click ok.

7) Close the popup and in the sheet, select any cell. Search for or type in function starting with "Faker_"...

If a set of functions starting with "Faker_" exists, you have installed the add-in successfully.

<h2>Functions</h2>

The following Excel functions have been implemented in the current Faker.NET release. Others are being added as an ongoing effort.

<ul>
<li>Faker_FirstName() - Returns a randomly selected First Name</li>
<li>Faker_Surname() - Returns a randomly selected Surname</li>
<li>Faker_Fullname() - Returns a randomly selected Full Name</li>
<li>Faker_StreetAddress() - Returns a randomly selected Street Address</li>
<li>Faker_StreetAddress(bool) - Returns a randomly selected StreetAddress. Pass parameter True if a secondary address to be returned.</li>
<li>Faker_StreetSuffix() - Returns a randomly selected Street Suffix</li>
<li>Faker_Country() - Returns a randomly selected Country</li>
<li>Faker_AusTown() - Returns a randomly selected Australian town or city</li>
<li>Faker_AusState() - Returns a randomly selected Australian state (full name)</li>
<li>Faker_AusPostcode() - Returns a randomly created Australian postcode (defined as four random digits)</li>
<li>Faker_AusPhoneNumber() - Returns a randomly created Australian phone/mobile number of various formats</li>
<li>Faker_AusUniversity() - Returns a randomly created Australian university of various formats</li>
<li>Faker_USState() - Returns a randomly selected US State.</li>
<li>Faker_USStateAbbr() - Returns a randomly selected, abbreviated US State.</li>
<li>Faker_USZipCode() - Returns a randomly selected US Zip Code.</li>
<li>Faker_UKCounty() - Returns a randomly selected UK County.</li>
<li>Faker_UKCountry() - Returns a randomly selected country in the UK.</li>
<li>Faker_UKPostCode() - Returns a randomly selected UK Postcode.</li>
<li>Faker_UKCompanyName() - Returns a randomly selected Company Name.</li>
<li>Faker_UKCompanyCatchphrase() - Returns a randomly selected Company Catchphrase.</li>
<li>Faker_Email() - Returns a randomly selected Email.</li>
<li>Faker_Phone() - Returns a randomly selected (US Format) Phone.</li>
<li>Faker_UserName() - Returns a randomly selected internet username.</li>
<li>Faker_CreditCard() - Returns a credit card number. Enter in types VISA, MASTERCARD or DINERSCLUB. VISA is default.</li>
</ul>


