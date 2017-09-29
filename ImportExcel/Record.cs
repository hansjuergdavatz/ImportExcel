using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
  public class DBField
  {
    public string Columen { get; set; }
    public string ColumenOrig { get; set; }
    public string TableName { get; set; }
    public bool ReadOny { get; set; }

    public DBField(string columen, string tableName, string columnOrig, bool readOnly)
    {
      Columen = columen;
      TableName = tableName;
      ColumenOrig = columnOrig;
      ReadOny = readOnly;
    }
  }

  public class DBFieldList : List<DBField>
  {
    public DBFieldList() { }

  }
}
