using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
  static class Extensions
  {
    public static string GetColumn(this DataRow Row, int Ordinal)
    {
      return Row.Table.Columns[Ordinal].ColumnName;
    }
  }
}
