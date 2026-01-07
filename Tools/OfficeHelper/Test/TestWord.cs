using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    public class TestWord
    {
        public static Dictionary<string, string> GetIssueTitleBlock()
        {
            string lstrDocumentNo = "lstrDocumentNo";
            string lstrDate = System.DateTime.Now.ToShortDateString();
            string lstrProjectName = "lstrProjectName";
            string lstrProjectCode = "lstrProjectCode";
            string lstrCustomerName = "lstrCustomerName";
            string lstrDepartment = "lstrDepartment";
            string lstrPartner = "lstrPartner";
            string lstrContract = "lstrContract";
            string lstrDesignPhase = "lstrDesignPhase";
            string lstrIssueType = "lstrIssue \r\n  Type";
            string lstrOtherIssueType = "lstrOtherIssueType";
            string lstrSealType = "lstrSeal \r\n Type";
            string lstrOtherSealType = "lstrOtherSealType";
            string lstrIssueDrawing = "lstrIssueDrawing";
            string lstrDesc = "lstrDesc";
            string lstrByCompany = "lstrByCompany";
            string lstrDiscipline = "lstrDiscipline";
            string lstrDrawingSize = "lstrDrawingSize";
            string lstrDrawingCount = "lstrDrawingCount";

         
            Dictionary<string, string> ldicTitleBlock = new Dictionary<string, string>{
                    {"txtDocumentNo",lstrDocumentNo},
                    {"txtDate",lstrDate},
                    {"txtDepartment",lstrDepartment},
                    {"txtPartner",lstrPartner},
                    {"txtContractNo",lstrContract},
                    {"txtProjectName",lstrProjectName},
                    {"txtProjectNo",lstrProjectCode},
                    {"txtCustomerName",lstrCustomerName},
                    {"txtIssueType",lstrIssueType},
                    {"txtOtherIssueType",lstrOtherIssueType},
                    {"txtSealType",lstrSealType},
                    {"txtOtherSealType",lstrOtherSealType},
                    {"txtIssueDrawingType",lstrIssueDrawing},
                    {"txtDesc",lstrDesc},
                    {"txtByCompany",lstrByCompany},
                    {"txtDiscipline",lstrDiscipline},
                    {"txtDrawingSize",lstrDrawingSize},
                    {"txtDrawingCount",lstrDrawingCount}
            };
            return ldicTitleBlock;
            

        }

        public static Dictionary<string, string> GetIssueForRefTitleBlock( )
        {
            string lstrDocumentNo = "lstrDocumentNo";
            string lstrDate = System.DateTime.Now.ToShortDateString();
            string lstrProjectName = "lstrProjectName";
            string lstrProjectCode = "lstrProjectCode";
            string lstrCustomerName = "lstrCustomerName";
            string lstrDepartment = "lstrDepartment";
            string lstrPartner = "lstrPartner";
            string lstrContract = "lstrContract";
            string lstrDesignPhase = "lstrDesignPhase";
            string lstrIssueType = "lstrIssue \r\n  Type";
            string lstrOtherIssueType = "lstrOtherIssueType";
            string lstrSealType = "lstrSeal \r\n Type";
            string lstrOtherSealType = "lstrOtherSealType";
            string lstrIssueDrawing = "lstrIssueDrawing";
            string lstrDesc = "lstrDesc";
            string lstrByCompany = "lstrByCompany";
            string lstrDiscipline = "lstrDiscipline";
            string lstrDrawingSize = "lstrDrawingSize";
            string lstrDrawingCount = "lstrDrawingCount";


            Dictionary<string, string> ldicTitleBlock = new Dictionary<string, string>{
                    {"txtDocumentNo",lstrDocumentNo},
                    {"txtDate",lstrDate},
                    {"txtDepartment",lstrDepartment},
                    {"txtPartner",lstrPartner},
                    {"txtContractNo",lstrContract},
                    {"txtProjectName",lstrProjectName},
                    {"txtProjectNo",lstrProjectCode},
                    {"txtCustomerName",lstrCustomerName},
                    {"txtIssueType",lstrIssueType},
                    {"txtOtherIssueType",lstrOtherIssueType},
                    {"txtSealType",lstrSealType},
                    {"txtOtherSealType",lstrOtherSealType},
                    {"txtIssueDrawingType",lstrIssueDrawing},
                    {"txtDesc",lstrDesc},
                    {"txtByCompany",lstrByCompany},
                    {"txtDiscipline",lstrDiscipline},
                    {"txtDrawingSize",lstrDrawingSize},
                    {"txtDrawingCount",lstrDrawingCount}
            };
            return ldicTitleBlock;
            

        }

    }
}
