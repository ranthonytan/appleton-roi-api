using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AppletonEmailAPI.Models
{
    public class EmailReport
    {
        public CustomerDetails Customer { get; set; }
        public ProjectDetails Project { get; set; }
        public Savings Savings { get; set; }
        public EnvironmentalImpact EnvironmentalImpact { get; set; }
        public bool isProposal { get; set; }//isBCCAllowed
        public bool isBCCAllowed { get; set; }
        //EnvironmentalImpact
    }

    public class CustomerDetails
    {
        public string CustomerName { get; set; }
        public string EmailAddress { get; set; }
        public string PhoneNumber { get; set; }
        public string CompanyName { get; set; }
        public string CompanyAddress { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public string PostalCode { get; set; }
    }
    public class ProjectDetails
    {
        public string ProjectName { get; set; }
        public string InstallationType { get; set; }
        public string Industry { get; set; }
        public string Currency { get; set; }
        public string EnergyCost { get; set; }
        public string TimePeriod { get; set; }
    }

    public class EnvironmentalImpact
    {
        public string RedEnergy { get; set; }
        public string SavedElectricity { get; set; }
        public string SavedTree { get; set; }
        public string CoalEmissionMetricTon { get; set; }
        public string CoalEmissionPound { get; set; }
        public string Car { get; set; }
        public string CO2MetricTon { get; set; }
        public string CO2Pound { get; set; }
    }
    public class Savings
    {
        public SavingDetails InitialInvestment { get; set; }
        public SavingDetails MaintenanceCosts { get; set; }
        public SavingDetails EnergyCosts { get; set; }
        public SavingDetails TotalCosts { get; set; }
        public ReportTotalSavings TotalSavings { get; set; }
        public ReportTotalSavings InitialNetInvest { get; set; }
        public ReportTotalSavings ROI { get; set; }
        public ReportTotalSavings AvgSaving { get; set; }
        public ReportTotalSavings PaybackPeriod { get; set; }
    }
    public class SavingDetails
    {
        public string ExistingLightingSystem { get; set; }
        public string ExistingLightingSystemPercentage { get; set; }
        public string AppletonLEDProposal { get; set; }
        public string AppletonLEDProposalPercentage { get; set; }
        public string AlternativeProposal { get; set; }
        public string AlternativeProposalPercentage { get; set; }
        public string ExistingLightingSystemString { get; set; }
        public string ExistingLightingSystemPercentageString { get; set; }
        public string AppletonLEDProposalString { get; set; }
        public string AppletonLEDProposalPercentageString { get; set; }
        public string AlternativeProposalString { get; set; }
        public string AlternativeProposalPercentageString { get; set; }
    }

    public class ReportTotalSavings
    {
        public string AppletonTotalSaving { get; set; }
        public string ProposalTotalSaving { get; set; }

    }

}