using System;
namespace PantheonProject
{
	public class CodesDescription
	{
		String codes;
		String description;
		public String getCodes()
		{
			return codes;
		}
		public void setCodes(String codes)
		{
			this.codes = codes;
		}
		public String getDescription()
		{
			return description;
		}
		public void setDescription(String description)
		{
			this.description = description;
		}

#pragma warning disable CS0114 // Member hides inherited member; missing override keyword
		public string ToString()
#pragma warning restore CS0114 // Member hides inherited member; missing override keyword
		{
			return "CodesDescription [codes=" + codes + ", description="
					+ description + "]";
		}
	}
}

