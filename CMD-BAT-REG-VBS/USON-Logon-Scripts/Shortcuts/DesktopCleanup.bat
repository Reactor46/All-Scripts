IF EXIST C:\DesktopCleanup.txt goto end else goto clean 
:clean
cd c:\
c:\ > DesktopCleanup.txt
del "%userprofile%\Desktop\Centennial Hills Radiolgy.url"
del "%userprofile%\Desktop\Desert Radiologists.url"
del "%userprofile%\Desktop\Desert Springs.url"
del "%userprofile%\Desktop\EMDAT.url"
del "%userprofile%\Desktop\Health Plan of Nevada, Inc. - HPN.url"
del "%userprofile%\Desktop\Lab Corp..url"
del "%userprofile%\Desktop\LMC.url"
del "%userprofile%\Desktop\Nevada Medicaid.url"
del "%userprofile%\Desktop\OPUS Valley Health Care.url"
del "%userprofile%\Desktop\Quest Lab (Care 360).url"
del "%userprofile%\Desktop\st.rose power chart.url"
del "%userprofile%\Desktop\Steinberg Diagnostics Physician Access Helpful Links.url"
del "%userprofile%\Desktop\sunrise.url"
del "%userprofile%\Desktop\UMC (Web Connect).url"
del "%userprofile%\Desktop\UMC Medical (eSign).url"
del "%userprofile%\Desktop\UMC PACs System.url"
del "%userprofile%\Desktop\West Valley Imaging.url"
:end




 