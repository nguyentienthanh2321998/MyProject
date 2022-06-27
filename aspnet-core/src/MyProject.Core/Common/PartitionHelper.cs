using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyProject.Common
{
    public class PartitionHelper
    {
        public PartitionHelper(int? pageIndex, int? pageSize)
        {
            ExecutePagingInformation(pageIndex, pageSize);
        }

        private int PageIndexDefault { get; } = 0;
        private int PageSizeMin { get; } = 10;
        private int PageSizeMax { get; } = 100;

        public int PageIndex { get; set; }
        public int PageSize { get; set; }
        public int From { get; set; }


        private void ExecutePagingInformation(int? pageIndex, int? pageSize)
        {
            PageIndex = !pageIndex.HasValue ? PageIndexDefault : pageIndex.Value;
            PageSize = !pageSize.HasValue ? PageSizeMin : pageSize.Value;

            if (PageIndex < PageIndexDefault) PageIndex = PageIndexDefault;

            if (PageSize > PageSizeMax) PageSize = PageSizeMax;

            if (PageSize < PageSizeMin) PageSize = PageSizeMin;

            From = PageIndex * PageSize;
        }
    }
}
