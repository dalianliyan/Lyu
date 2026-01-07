namespace MarvellousWorks.PracticalPattern.Concept.Generics
{
    public interface IOrganization { }

    public abstract class UserBase<TKey, TOrganization>
        where TOrganization : class, IOrganization, new()
    {
        // 用于约束属性方法
        public abstract TOrganization Organization { get;}      

        //  用于约束方法
        public abstract void Promotion(TOrganization newOrganization);

        //  用于约束委托
        delegate TOrganization OrganizationChangedHandler(); 
    }
}
