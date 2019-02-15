using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace TimeLiveRegression.org.com.elements
{
    public class Elements
    {
        Dictionary<string, By> hash = new Dictionary<string, By>();
        public Elements()
        {

            /*  //Login Page
              hash.Add("edtUserName", By.Id("Login:LoginScreen:LoginDV:username-inputEl"));
              hash.Add("edtPassword", By.Id("Login:LoginScreen:LoginDV:password-inputEl"));
              hash.Add("eleLogIn", By.Id("Login:LoginScreen:LoginDV:submit-btnInnerEl"));

              //Search Form
              hash.Add("eleClaim", By.Id("TabBar:ClaimTab-btnInnerEl"));
              hash.Add("eleClaimAction", By.Id("Claim:ClaimMenuActions-btnIconEl"));
              hash.Add("eleNote", By.Id("Claim:ClaimMenuActions:ClaimNewOtherMenuItemSet:ClaimMenuActions_NewOther:ClaimMenuActions_NewNote-textEl"));
              hash.Add("lstTopic", By.Id("NewNoteWorksheet:NewNoteScreen:NoteDetailDV:Topic-inputEl"));
              hash.Add("lstRelatedTo", By.Id("NewNoteWorksheet:NewNoteScreen:NoteDetailDV:RelatedTo-inputEl"));
              hash.Add("edtText", By.Id("NewNoteWorksheet:NewNoteScreen:NoteDetailDV:Body-inputEl"));
              hash.Add("eleUpdate", By.Id("NewNoteWorksheet:NewNoteScreen:Update-btnInnerEl"));
              hash.Add("eleClear", By.Id("WebMessageWorksheet:WebMessageWorksheetScreen:WebMessageWorksheet_ClearButton-btnInnerEl"));

              //LogOut Page
              hash.Add("elePreferences", By.Id(":TabLinkMenuButton-btnIconEl"));
              hash.Add("eleLogOut", By.Id(":TabBar:LogoutTabBarLink-textEl"));*/

            //Login Page
            hash.Add("edtUserName", By.Id("username"));
            hash.Add("edtPassword", By.Id("password"));
            hash.Add("btnLogin", By.Id("login"));

            //Search Form
            hash.Add("lstLocation", By.Id("location"));
            hash.Add("lstHotels", By.Id("hotels"));
            hash.Add("lstRoomType", By.Id("room_type"));
            hash.Add("lstNoOfRooms", By.Id("room_nos"));
            hash.Add("edtDatePickIn", By.Id("datepick_in"));
            hash.Add("edtDatePickOut", By.Id("datepick_out"));
            hash.Add("lstAdult_Rom", By.Id("adult_room"));
            hash.Add("lstChildRoom", By.Id("child_room"));
            hash.Add("eleSubmit", By.Id("Submit"));


            //LogOut Page
            hash.Add("btnLogOut", By.XPath("/html/body/table[2]/tbody/tr[1]/td[2]/a[4]"));

        }
        public By getObject(string element)
        {
            By returnValue = null;
            if (hash.ContainsKey(element))
            {
                //returnValue = hash.TryGetValue(element);
                hash.TryGetValue(element, out returnValue);
            }
            return returnValue;
        }
    }
}
