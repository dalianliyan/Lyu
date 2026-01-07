using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace SGJ.Utilities
{

    public class EDRIDrawingList
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public string ProjectName { get; set; }

        public string ProjectNo { get; set; }

        public Dictionary<string, EDRIDiscipline> Disciplines { get; set; }

        public EDRIDrawingList()
        {
            Disciplines = new Dictionary<string, EDRIDiscipline>();
        }
    }

    public class EDRIDiscipline
    {
        public string Name { get; set; }

        public string Description { get; set; }

        public Dictionary<string, EDRIBuilding> Buildings { get; set; }

        public EDRIDiscipline()
        {
            Buildings = new Dictionary<string, EDRIBuilding>();
        }
    }

    public class EDRIBuilding
    {

        public string Name { get; set; }

        public string Description { get; set; }

        public List<EDRIDesignDocument> Documents { get; set; }

        public EDRIBuilding()
        {
            Documents = new List<EDRIDesignDocument>();
        }
    }

    public class EDRIDesignDocument
    {
        #region Constructor
        public EDRIDesignDocument(string name, string desc, string discipline,string building)
        {
            this.DrawingNo = name;
            this.Description = desc;
            this.DisciplineName = discipline;
            this.BuildingName = building;
        }
        #endregion

        #region Properties

        public string DrawingNo { get; set; }

        public string NameEn { get; set; }

        public string Description { get; set; }

        public string Size { get; set; }

        public string Remarks { get; set; }

        /// <summary>
        /// MajorRevision
        /// </summary>
        public string R0No { get; set; }

        /// <summary>
        /// MinorRevision
        /// </summary>
        public string R1No { get; set; }

        public string Designer { get; set; }

        public string Checker { get; set; }

        public string Reviewer { get; set; }

        public string Approver { get; set; }

        public string BuildingName { get; set; }

        public string DisciplineName { get; set; }
        #endregion

    }

}
