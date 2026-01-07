using System.Collections.Generic;
namespace MarvellousWorks.PracticalPattern.Concept.Tests
{
    public class Education{}
    public class Experience{} 


    public interface IPerson
    {
        void Promote(string newTitle);
        string Department { get; set; }
        IList<Education> EducationHistory { get; set; }
        IList<Experience> WorkingExperience { get; set; }
    }
}
