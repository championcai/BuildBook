<%@ Page language="c#" Debug="true"%>
<%
// ---数据初始化----------------------------------------------------------------------------------------------------------------------
        Response.ContentType = "text/xml";
        Response.Charset = "utf-8";
        Response.AppendHeader("Cache-Control", "max-age=1, must-revalidate");
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        Response.Write("<Grid>");

        string SQLserver="server=.;pwd=sa;UID=sa;database=CTS";
        string SQLtable="BuildBook";
        System.Data.SqlClient.SqlConnection Conn=new System.Data.SqlClient.SqlConnection(SQLserver);Conn.Open();
        System.Data.SqlClient.SqlCommand Cmd = Conn.CreateCommand();
        System.Data.SqlClient.SqlDataReader R;

        string Mode = Request["Mode"]== null ? "浏览" : Request["Mode"];
        string DelMode =Request["DelMode"];
        string User = System.Web.HttpContext.Current.Request.Cookies["KEACookies"]["KEAuser"].ToUpper();
        string Dept = System.Web.HttpContext.Current.Request.Cookies["KEACookies"]["KEAdept"].ToUpper();
        string Role = System.Web.HttpContext.Current.Request.Cookies["KEACookies"]["KEArole"].ToUpper();

        string currentTime=System.DateTime.Now.ToString("d");


// --- 将Treegrid的变化内容写回到数据库中去-------------------------------------------------------------------------------------------
       string XML = Request["TGData"];
       if (XML != "" && XML != null)
       {
          if(User=="") Response.Write ("<IO Result='-1' Message='请登录KEATM!'/></Grid>");
          else
          {
             System.Xml.XmlDocument X = new System.Xml.XmlDocument();
             X.LoadXml(HttpUtility.HtmlDecode(XML));
             System.Xml.XmlNodeList Ch = X.GetElementsByTagName("Changes"); //检索变更的XML内容

             if (Ch.Count > 0) foreach (System.Xml.XmlElement I in Ch[0])  //逐条检查变更的XML内容
             {
                string[] ids = I.GetAttribute("id").Split("$".ToCharArray());

                     if (I.GetAttribute("Deleted")=="1") //销毁
                      {
                           Cmd.CommandText = "DELETE FROM "+SQLtable+" WHERE rowid=" + ids[2];
                           Cmd.ExecuteNonQuery();
                        //   Response.Write ("<IO Result='0' Message='销毁成功'/></Grid>");
                        Response.Write ("</Grid>");
                      }
                      if (I.GetAttribute("Added")=="1") //添加
                      {
                           Cmd.CommandText = "INSERT INTO "+SQLtable+"(Code,Name,NameEN,Type,属性1,属性2,属性3,属性4,属性5,属性6,属性7,属性8,属性9,Description,Comments,History,Owner,Department,Status,At,Mark) VALUES ('";
                           Cmd.CommandText+= I.GetAttribute("CCode")+ "','" ;
                           Cmd.CommandText+= I.GetAttribute("CName").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','" ;
                           Cmd.CommandText+= I.GetAttribute("CNameEN").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','" ;
                           Cmd.CommandText+= I.GetAttribute("CType")+ "','" ;
                           Cmd.CommandText+= I.GetAttribute("C属性1").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性2").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性3").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性4").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性5").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','" ;
                           Cmd.CommandText+= I.GetAttribute("C属性6").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性7").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性8").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("C属性9").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','"  ;
                           Cmd.CommandText+= I.GetAttribute("CDescription").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;").Replace("\"","&quot;") + "','" ;
                           Cmd.CommandText+= I.GetAttribute("CComments").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "','";
                           Cmd.CommandText+= User+"["+currentTime+"][新增]:" + I.GetAttribute("CComments")+ "','";
                           Cmd.CommandText+= I.GetAttribute("COwner") +User+ "','";
                           Cmd.CommandText+= I.GetAttribute("CDepartment") +Dept+ "','";
                           Cmd.CommandText+= I.GetAttribute("CStatus") + "','";
                           Cmd.CommandText+= currentTime + "','";
                           Cmd.CommandText+= "A')";
                           Cmd.ExecuteNonQuery();
                        //   Response.Write ("<IO Result='0' Message='添加成功'/></Grid>");
                        Response.Write ("</Grid>");
                      }
                      if (I.GetAttribute("Changed")=="1") //修改和伪删除
                      {
                           string Str ="";
                           if(I.HasAttribute("CCode"))  Str += "Code='" +I.GetAttribute("CCode").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("CName")) Str += "Name='" +I.GetAttribute("CName").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("CNameEN")) Str += "NameEN='" +I.GetAttribute("CNameEN").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("CType")) Str += "Type='" +I.GetAttribute("CType").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性1")) Str += "属性1='" + I.GetAttribute("C属性1").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性2")) Str += "属性2='" + I.GetAttribute("C属性2").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性3")) Str += "属性3='" + I.GetAttribute("C属性3").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性4")) Str += "属性4='" + I.GetAttribute("C属性4").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性5")) Str += "属性5='" + I.GetAttribute("C属性5").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性6")) Str += "属性6='" + I.GetAttribute("C属性6").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性7")) Str += "属性7='" + I.GetAttribute("C属性7").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性8")) Str += "属性8='" + I.GetAttribute("C属性8").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("C属性9")) Str += "属性9='" + I.GetAttribute("C属性9").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           if(I.HasAttribute("CDescription")) Str += "Description='" +I.GetAttribute("CDescription").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;").Replace("\"","&quot;") + "',";
                           if(I.HasAttribute("CComments")) Str += "Comments='" +I.GetAttribute("CComments").Replace("&","&amp;").Replace("'","&apos;").Replace("<","&lt;") + "',";
                           Str += "History='"+User+"["+currentTime+"]:" + I.GetAttribute("CComments") + "\r\n"+ I.GetAttribute("CHistory")+"',";
                           if(I.HasAttribute("COwner")) Str += "Owner='" + I.GetAttribute("COwner") + "',";
                           if(Role.Contains("KEATM_ADMIN")&&I.HasAttribute("CDepartment")) Str += "Department='" + I.GetAttribute("CDepartment") + "',";                            //只有角色为KEATM_ADMIN或KEATM_COORDINATOR时，才能修改责任部门
                           if(I.HasAttribute("CStatus")) Str += "Status='" + I.GetAttribute("CStatus") + "',";
                           Str += "At='" + currentTime+ "',";
                           if(I.HasAttribute("CMark")) Str += "Mark='" + I.GetAttribute("CMark") + "',";                   //通过设置Mark为‘D’达到伪删除的目的
                           if(Str != "") {
                               Cmd.CommandText = "UPDATE "+SQLtable+" SET " + Str.TrimEnd(",".ToCharArray()) + " WHERE rowid="+ ids[2];
                               Cmd.ExecuteNonQuery();
                               if(I.GetAttribute("CMark")=="D") Response.Write ("<IO Result='0' Message='删除成功'/></Grid>");
                        //       else Response.Write ("<IO Result='0' Message='修改成功'/></Grid>");
                               else Response.Write ("</Grid>");
                           }
                      }
             }
          }
       }
// －－－－－从数据库读取数据到Treegrid里面来－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
       else
       {
          Cmd.CommandText =  "SELECT * FROM "+SQLtable;
           if (Role.Contains("KEATM_ADMIN")) Cmd.CommandText+="  order by Code ";
          else Cmd.CommandText+=" where Owner='"+User+"' order by Code ";
          R = Cmd.ExecuteReader();
          string Str = "";
          Str += "<Body><B>";
          while(R.Read()){
                   Str += "<I  Def='MAIN'  "
                      + " Ident='" + R["rowid"].ToString() + "'"
                      + " CCode='" + R["Code"].ToString()+ "'"
                      + " CName='" + R["Name"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " CNameEN='" + R["NameEN"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " CType='" + R["Type"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性1='" + R["属性1"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性2='" + R["属性2"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性3='" + R["属性3"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性4='" + R["属性4"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性5='" + R["属性5"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性6='" + R["属性6"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性7='" + R["属性7"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性8='" + R["属性8"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " C属性9='" + R["属性9"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " CDescription='" + R["Description"].ToString() + "'"
                      + " CComments='" + R["Comments"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " CHistory='" + R["History"].ToString().Replace( "&amp;","&").Replace( "&apos;","’").Replace( "&lt;","<") + "'"
                      + " COwner='" + R["Owner"].ToString() + "'"
                      + " CDepartment='" + R["Department"].ToString() + "'"
                      + " CStatus='" + R["Status"].ToString() + "'"
                      + " CAt='" + R["At"].ToString() + "'"
                      + " CMark='" + R["Mark"].ToString() + "'"
                      + "></I>";
              }

          Str += "</B></Body></Grid>";
          Response.Write(Str);
          R.Close();
       }
       Conn.Close();
%>