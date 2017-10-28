using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Faker;

namespace FakerNETExcel
{
    /// <summary>
    /// Maps to all Excel Functions
    /// </summary>
    public static class FakerExcelFunctions
    {
        [ExcelFunction(Description = "Returns a randomly selected First Name")]
        public static string Faker_FirstName()
        {
            return Faker.Name.GetFirstName();
        }

        [ExcelFunction(Description = "Returns a randomly selected Surname")]
        public static string Faker_Surname()
        {
            return Faker.Name.GetLastName();
        }

        [ExcelFunction(Description = "Returns a randomly selected Full Name")]
        public static string Faker_Fullname()
        {
            return Faker.Name.GetName();
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetAddress")]
        public static string Faker_StreetAddress()
        {
            return Faker.Address.GetStreetAddress();
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetAddress. Pass parameter True if a secondary address to be returned.")]
        public static string Faker_StreetAddress(bool IncludeSecondary)
        {
            return Faker.Address.GetStreetAddress(IncludeSecondary);
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetName.")]
        public static string Faker_StreetName()
        {
            return Faker.Address.GetStreetName();
        }

        [ExcelFunction(Description = "Returns a randomly selected Street Suffix.")]
        public static string Faker_StreetSuffix()
        {
            return Faker.Address.GetStreetSuffix();
        }

        [ExcelFunction(Description = "Returns a randomly selected City.")]
        public static string Faker_City()
        {
            return Faker.Address.GetCity();
        }

        [ExcelFunction(Description = "Returns a randomly selected Country.")]
        public static string Faker_Country()
        {
            return Faker.Address.GetCountry();
        }

        [ExcelFunction(Description = "Returns a randomly selected AU Town or City.")]
        public static string Faker_AusTown()
        {
            return Faker.Address.GetAusTown();
        }

        [ExcelFunction(Description = "Returns a randomly selected AU University.")]
        public static string Faker_AusUniversity()
        {
            return Faker.Education.GetAusUni();
        }

        [ExcelFunction(Description = "Returns a randomly selected AU State.")]
        public static string Faker_AusState()
        {
            return Faker.Address.GetAusState();
        }

        [ExcelFunction(Description = "Returns a randomly selected AU Postcode.")]
        public static string Faker_AusPostcode()
        {
            return Faker.Address.GetAusPostcode();
        }

        [ExcelFunction(Description = "Returns a randomly selected AU Phone Number of various types")]
        public static string Faker_AusPhoneNumber()
        {
            return Faker.PhoneNumber.GetAusPhoneNumber();
        }

        [ExcelFunction(Description = "Returns a randomly selected US State.")]
        public static string Faker_USState()
        {
            return Faker.Address.GetUSState();
        }

        [ExcelFunction(Description = "Returns a randomly selected, abbreviated US State.")]
        public static string Faker_USStateAbbr()
        {
            return Faker.Address.GetUSStateAbbr();
        }

        [ExcelFunction(Description = "Returns a randomly selected US Zip Code.")]
        public static string Faker_USZipCode()
        {
            return Faker.Address.GetZipCode();
        }

        [ExcelFunction(Description = "Returns a randomly selected UK County.")]
        public static string Faker_UKCounty()
        {
            return Faker.Address.GetUKCounty();
        }

        [ExcelFunction(Description = "Returns a randomly selected country in the UK.")]
        public static string Faker_UKCountry()
        {
            return Faker.Address.GetUKCountry();
        }

        [ExcelFunction(Description = "Returns a randomly selected UK postcode.")]
        public static string Faker_UKPostcode()
        {
            return Faker.Address.GetUKPostcode();
        }

        [ExcelFunction(Description = "Returns a randomly selected US Phone Number of various types")]
        public static string Faker_USPhoneNumber()
        {
            return Faker.PhoneNumber.GetUSPhoneNumber();
        }

        [ExcelFunction(Description = "Returns a randomly selected company name.")]
        public static string Faker_CompanyName()
        {
            return Faker.Company.GetName();
        }

        [ExcelFunction(Description = "Returns a randomly selected company catchphrase.")]
        public static string Faker_CompanyCatchphrase()
        {
            return Faker.Company.GetCatchPhrase();
        }

        [ExcelFunction(Description = "Returns a randomly created email. To base the email on a name, add name as parameter.")]
        public static string Faker_Email(string name)
        {
            return Faker.Internet.GetEmail(name);
        }

        [ExcelFunction(Description = "Returns a randomly created email. To base the email on a name, add name as parameter.")]
        public static string Faker_Email()
        {
            return Faker_Email();
        }

        [ExcelFunction(Description = "Returns a randomly created (US format) Phone Number.")]
        public static string Faker_Phone()
        {
            return Faker.PhoneNumber.GetUSPhoneNumber();
        }

        [ExcelFunction(Description = "Returns a randomly created internet username. To base it on a name, add name as parameter.")]
        public static string Faker_UserName(string name = null)
        {
            return Faker.Internet.GetUserName(name);
        }

        [ExcelFunction(Description = "Returns a randomly created internet username. To base it on a name, add name as parameter.")]
        public static string Faker_UserName()
        {
            return Faker_UserName();
        }


    }
}
