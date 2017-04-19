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
            return Faker.Name.First();
        }

        [ExcelFunction(Description = "Returns a randomly selected Surname")]
        public static string Faker_Surname()
        {
            return Faker.Name.Last();
        }

        [ExcelFunction(Description = "Returns a randomly selected Full Name")]
        public static string Faker_Fullname()
        {
            return Faker.Name.FullName();
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetAddress")]
        public static string Faker_StreetAddress()
        {
            return Faker.Address.StreetAddress(false);
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetAddress. Pass parameter True if a secondary address to be returned.")]
        public static string Faker_StreetAddress(bool IncludeSecondary)
        {
            return Faker.Address.StreetAddress(IncludeSecondary);
        }

        [ExcelFunction(Description = "Returns a randomly selected StreetName.")]
        public static string Faker_StreetName()
        {
            return Faker.Address.StreetName();
        }

        [ExcelFunction(Description = "Returns a randomly selected Street Suffix.")]
        public static string Faker_StreetSuffix()
        {
            return Faker.Address.StreetSuffix();
        }

        [ExcelFunction(Description = "Returns a randomly selected City.")]
        public static string Faker_City()
        {
            return Faker.Address.City();
        }

        [ExcelFunction(Description = "Returns a randomly selected Country.")]
        public static string Faker_Country()
        {
            return Faker.Address.Country();
        }

        [ExcelFunction(Description = "Returns a randomly selected US State.")]
        public static string Faker_USState()
        {
            return Faker.Address.UsState();
        }

        [ExcelFunction(Description = "Returns a randomly selected, abbreviated US State.")]
        public static string Faker_USStateAbbr()
        {
            return Faker.Address.UsStateAbbr();
        }

        [ExcelFunction(Description = "Returns a randomly selected US Zip Code.")]
        public static string Faker_USZipCode()
        {
            return Faker.Address.UsStateAbbr();
        }

        [ExcelFunction(Description = "Returns a randomly selected UK County.")]
        public static string Faker_UKCounty()
        {
            return Faker.Address.UkCounty();
        }

        [ExcelFunction(Description = "Returns a randomly selected country in the UK.")]
        public static string Faker_UKCountry()
        {
            return Faker.Address.UkCountry();
        }

        [ExcelFunction(Description = "Returns a randomly selected UK postcode.")]
        public static string Faker_UKPostcode()
        {
            return Faker.Address.UkPostCode();
        }

        [ExcelFunction(Description = "Returns a randomly selected company name.")]
        public static string Faker_CompanyName()
        {
            return Faker.Company.Name();
        }

        [ExcelFunction(Description = "Returns a randomly selected company catchphrase.")]
        public static string Faker_CompanyCatchphrase()
        {
            return Faker.Company.CatchPhrase();
        }

        [ExcelFunction(Description = "Returns a randomly created email. To base the email on a name, add name as parameter.")]
        public static string Faker_Email(string name)
        {
            return Faker.Internet.Email(name);
        }

        [ExcelFunction(Description = "Returns a randomly created email. To base the email on a name, add name as parameter.")]
        public static string Faker_Email()
        {
            return Faker.Internet.Email();
        }

        [ExcelFunction(Description = "Returns a randomly created (US format) Phone Number.")]
        public static string Faker_Phone()
        {
            return Faker.Phone.Number();
        }

        [ExcelFunction(Description = "Returns a randomly created internet username. To base it on a name, add name as parameter.")]
        public static string Faker_UserName(string name)
        {
            return Faker.Internet.UserName(name);
        }

        [ExcelFunction(Description = "Returns a randomly created internet username. To base it on a name, add name as parameter.")]
        public static string Faker_UserName()
        {
            return Faker.Internet.UserName();
        }


    }
}
