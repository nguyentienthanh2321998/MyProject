using Abp.Domain.Entities.Auditing;

namespace MyProject.Models
{
    public abstract class AuditedSoftDeletableEntity<TPrimaryKey> : AuditedEntity<TPrimaryKey>
    {
        public bool IsActive { get; set; }
    }
}