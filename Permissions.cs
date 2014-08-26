using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MakeUpExam
{
	class Permissions
	{
		public static bool 補考學生清單權限
		{
			get
			{
				return FISCA.Permission.UserAcl.Current[補考學生清單].Executable;
			}
		}

		public static string 補考學生清單 = "MakeUpExam-{614B3F80-0402-44AA-994F-DC7EBF55700C}";
	}
}
