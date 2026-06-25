//objectcomments Generated Application Object
forward
global type pypbexample from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global variables
String gs_appdir   // directorio de trabajo inicial (donde vive bin.pypb.appeon\)
end variables

global type pypbexample from application
string appname = "pypbexample"

string themepath = "C:\Program Files (x86)\Appeon\PowerBuilder 25.0\IDE\theme"
string themename = "Do Not Use Themes"
boolean nativepdfvalid = false
boolean nativepdfincludecustomfont = false
string nativepdfappname = ""
long richtextedittype = 5
long richtexteditx64type = 5
long richtexteditversion = 3
string richtexteditkey = ""
string appicon = ""
string appruntimeversion = "25.1.0.6430"
boolean manualsession = false
boolean unsupportedapierror = false
boolean ultrafast = false
boolean bignoreservercertificate = false
uint ignoreservercertificate = 0
long webview2distribution = 0
boolean webview2checkx86 = false
boolean webview2checkx64 = false
string webview2url = "https://developer.microsoft.com/en-us/microsoft-edge/webview2/"
end type
global pypbexample pypbexample

type prototypes
// Funcion para abrir archivos externos con el programa asociado.
FUNCTION long ShellExecute(ulong hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, ulong nShowCmd) LIBRARY "SHELL32.DLL" alias for "ShellExecuteW"

//Funcion para tomar el directorio de la aplicacion  -64Bits 
FUNCTION	uLong	GetModuleFileName ( uLong lhModule, ref string sFileName, ulong nSize )  LIBRARY "Kernel32.dll" ALIAS FOR "GetModuleFileNameW"
end prototypes

on pypbexample.create
appname="pypbexample"
message=create message
sqlca=create transaction
sqlda=create dynamicdescriptionarea
sqlsa=create dynamicstagingarea
error=create error
end on

on pypbexample.destroy
destroy(sqlca)
destroy(sqlda)
destroy(sqlsa)
destroy(error)
destroy(message)
end on

event open;String ls_Path, ls_appName="pypbexample"
ulong lul_handle


ls_Path = space(1024)
SetNull(lul_handle)
GetModuleFilename(lul_handle, ls_Path, len(ls_Path))

if right(UPPER(ls_path), 7)="250.EXE" or right(UPPER(ls_path), 7)="X64.EXE" then
	ls_path="C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\"+ls_appName+".exe"
end if

gs_appdir=left(ls_path, len(ls_path) - (len(ls_appName) + 5))

// Pre-cargar PyPb al arranque: su assembly .NET debe inicializarse ANTES que el
// export XLSX nativo (si no, el orden nativo->Python falla al cargar bin.pypb.appeon).
n_cst_pyton lnv_pyinit
lnv_pyinit = CREATE n_cst_pyton
lnv_pyinit.of_init(gs_appdir + "\python.runtime\python313.dll")
DESTROY lnv_pyinit

open(w_main)
end event

