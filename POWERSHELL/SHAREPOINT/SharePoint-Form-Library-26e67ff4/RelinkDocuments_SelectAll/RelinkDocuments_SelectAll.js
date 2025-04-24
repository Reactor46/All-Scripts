var form = document.aspnetForm;
for (var i=0; i<form.elements.length; i++)
{
if (form.elements[i].type == "checkbox") 
{
if (form.elements[i].id == 'chkRepair')
{
form.elements[i].checked = true;
}
}
}
