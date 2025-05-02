<%@ Assembly Name="KSC.SharePoint.PublicWebsite, Version=1.0.0.0, Culture=neutral, PublicKeyToken=be655621e85068cf" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ScheduleAnCOVIDAppointmentUserControl.ascx.cs" Inherits="KSC.SharePoint.PublicWebsite.Webparts.ScheduleAnCOVIDAppointmentUserControl" %>
<%@ Register Assembly="AjaxControlToolkit, Version=15.1.4.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e" Namespace="AjaxControlToolkit" TagPrefix="AjaxControlToolkit" %>

<asp:Panel runat="server" ID="PanelMainDisplay" CssClass="sa-webpart-panel">

        <!-- webpart update panel -->
        <div class="sa-webpart-container">

            <!-- title -->
            <h2>Schedule a COVID-19 Vaccine Appointment</h2>

            <!-- sub content-->
            <p>Directly schedule your free COVID-19 vaccine appointment. Appointments open to new and current patients 5 years and older.</p>

            <asp:UpdatePanel runat="server" ID="UpdatePanelControlContainer" ChildrenAsTriggers="true" UpdateMode="Always">

                <ContentTemplate>
                    <div class="sa-control-container">
                        <asp:Button runat="server" ID="ButtonDummy" style="display:none;" />
                        <asp:RadioButtonList runat="server" ID="RadioButtonListVisitTypes">
                            <asp:ListItem Value="708" Text="COVID-19 Vaccination" Selected="True"></asp:ListItem>
                        </asp:RadioButtonList>
                        <asp:TextBox runat="server" ID="TextBoxDateOfBirth" placeholder="Date of Birth ( mm/dd/yyyy )" CssClass="child-control"></asp:TextBox>
                        <asp:Button runat="server" ID="ButtonSearch" Text="search" CssClass="child-control" OnClick="ButtonSearch_Click"/>
                        <AjaxControlToolkit:ModalPopupExtender ID="ModalPopup" runat="server" TargetControlID="ButtonDummy" PopupControlID="PanelScreeningQuestions" BackgroundCssClass="modal-backdrop" Y="5"></AjaxControlToolkit:ModalPopupExtender>
                    </div><!--/end control-container -->
                </ContentTemplate>
                <Triggers>
                    <asp:PostBackTrigger ControlID="ButtonSearch" />
                </Triggers>
            </asp:UpdatePanel>

            <div><asp:Label runat="server" ID="LabelMessage"></asp:Label></div>

            <!-- bottom links -->
            <ul>
                <li><a href="https://www.kelsey-seybold.com/mykelseyonline">Patient<br />Log In</a></li>
                <li><a href="https://www.kelsey-seybold.com/health-information/covid-19/pages/covid-19-faq.aspx">COVID-19<br />Vaccine FAQ</a></li>
            </ul>

        </div><!--/end webpart-container -->

</asp:Panel>

<asp:Panel runat="server" ID="PanelScreeningQuestions" CssClass="questions-popup" style="display:none;">

    <asp:Button runat="server" ID="ButtonDummyButtonNo" style="display:none;" />

    <!-- modal panel -->
    <section class="modal-container">

        <div class="modal-body">

            <div id="modal-question" class="modal-question">
                <h2 class="modal-header">For the safety of our patients and employees, please review and accept the following COVID-19 disclaimer: </h2>

                <ul class="modal-ul">
                    <li>I understand that to schedule for the Pfizer vaccine I must be at least 5 years old.</li>
                    <li>I understand that to schedule for the Moderna or Jansen (J&J) vaccine I must be at least 18 years old.</li>
                    <li>I have not received any COVID vaccinations and am not currently scheduled for a COVID vaccination.</li>
                    <li>I have not had a fever or illness in the last 14 days.</li>
                    <li>I have not tested positive for COVID in the last 30 days.</li>
                    <li>By clicking "Accept," you acknowledge that you have read the COVID-19 Emergency Use Authorization fact sheets for Pfizer, Janssen (J&J) and Moderna below (clicking the links will open a new tab to the fact sheets):
                        <ul>
                            <li><a href="http://labeling.pfizer.com/ShowLabeling.aspx?id=14472&format=pdf" target="_blank">Pfizer COVID-19 Vaccine EUA</a></li>
                            <li><a href="https://www.mykelseyonline.com/MyChart/en-US/docs/jj_vaccine_info_sheet.pdf" target="_blank">Janssen (J&J) COVID-19 Vaccine EUA</a></li>
                            <li><a href="https://www.modernatx.com/covid19vaccine-eua/eua-fact-sheet-recipients.pdf" target="_blank">Moderna COVID-19 Vaccine EUA</a></li>
                        </ul>
                    </li>
                    <li>I do not have a history of Guillain Barre Syndrome.</li>
                    <li>I have not received any other vaccinations in the last 14 days.</li>
                    <li>I have not received convalescent plasma for SARS-CoV-2 (COVID-19) in the last 90 days.</li>
                    <li>I have not received monoclonal antibody infusions for SARS-CoV-2 (COVID-19) in the last 90 days.</li>
                    <li>Any minors receiving vaccines should be accompanied by a parent/guardian to provide permission.</li>
                    <li>I understand that protection against COVID-19 may not be effective until at least 7 days after the single dose administration for Janssen (J&J) and may not be effective until at least 7 days after the second dose for Pfizer or Moderna COVID-19 vaccine.</li>
                    <li>
                        By clicking the “Accept” button below I attest that I am at least 5 years old and attest to all items above. 
                        Once you have clicked “Accept” you will choose an available location, vaccine type and date/time for your appointment and subsequently provide your demographic and insurance information. 
                        The vaccine is provided at no cost to the patient. 
                        Vaccine providers can be reimbursed by your health insurance provider for the administration of the vaccine.  Available sites and vaccine types are limited.
                    </li>
                </ul>

                <div class="modal-button-container">
                    <asp:Button runat="server" ID="ButtonAccept" Text="Accept" OnClick="ButtonAccept_Click" />
                    <AjaxControlToolkit:ModalPopupExtender ID="ModalPopupResults" runat="server" TargetControlID="ButtonDummyButtonNo" PopupControlID="PanelResults" CancelControlID="ButtonResultsClose" BackgroundCssClass="modal-backdrop" y="20"></AjaxControlToolkit:ModalPopupExtender>
                </div>
            </div>

            <div id="modal-disclaimer" class="modal-disclaimer">
                <h2 class="modal-header">We apologize for the inconvenience, based on your symptoms, we are unable to schedule your in-person appointment.</h2>
                <div>
                    <p>We can still see you immediately! Please call our Contact Center at 713-442-0000 to continue scheduling your appointment, we are available 24/7 to assist you.</p>
                </div>
                <div class="modal-button-container">
                    <asp:Button runat="server" ID="ButtonClose" Text="Close" CausesValidation="false" UseSubmitBehavior="false" />
                </div>
            </div>

        </div>

    </section>

    <script type="text/javascript">

        function ToggleModalPanel() {

            var questionPanel = document.getElementById("modal-question");
            var disclaimerPanel = document.getElementById("modal-disclaimer");

            questionPanel.style.display = "none";
            disclaimerPanel.style.display = "block";

        }//end function

    </script>

</asp:Panel>

<asp:Panel runat="server" ID="PanelResults" CssClass="results-panel" style="display:none;">

    <!-- results -->
    <section class="results-modal-container">
        <div class="results-modal-body">
            <h2 class="results-modal-header">
                <asp:Label runat="server" ID="LabelResultsHeader" CssClass="results-modal-header-span"></asp:Label>
                <asp:Button runat="server" ID="ButtonResultsClose" Text="X" UseSubmitBehavior="false" CausesValidation="false" OnClick="ButtonResultsClose_Click" />
            </h2>
            <div class="results-modal-iframe-container">
                <asp:Literal runat="server" ID="LiteralSearchResults"></asp:Literal>
            </div>
        </div>
        
    </section>

</asp:Panel>
