//objectcomments  
forward
global type w_main from window
end type
type p_2 from picture within w_main
end type
type st_info from statictext within w_main
end type
type st_myversion from statictext within w_main
end type
type st_platform from statictext within w_main
end type
type r_2 from rectangle within w_main
end type
type cb_probar from commandbutton within w_main
end type
type st_pyversion from statictext within w_main
end type
type cb_excel from commandbutton within w_main
end type
type dw_1 from datawindow within w_main
end type
type cb_cargar from commandbutton within w_main
end type
type cb_xls_nat from commandbutton within w_main
end type
type cb_xls_py from commandbutton within w_main
end type
end forward

global type w_main from window
integer width = 6501
integer height = 2712
boolean titlebar = true
string title = "Pyton Example"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
string icon = "AppIcon!"
boolean center = true
p_2 p_2
st_info st_info
st_myversion st_myversion
st_platform st_platform
r_2 r_2
cb_probar cb_probar
st_pyversion st_pyversion
cb_excel cb_excel
dw_1 dw_1
cb_cargar cb_cargar
cb_xls_nat cb_xls_nat
cb_xls_py cb_xls_py
end type
global w_main w_main

type prototypes

end prototypes

type variables

end variables

forward prototypes
public subroutine wf_version ()
public subroutine wf_probar_python ()
public subroutine wf_probar_excel ()
public subroutine wf_cargar_facturas ()
public subroutine wf_excel_nativo ()
public subroutine wf_excel_python ()
end prototypes

public subroutine wf_version ();String ls_version, ls_platform
Integer li_rtn
Environment l_env

li_rtn = GetEnvironment(l_env)

IF li_rtn <> 1 THEN 
	ls_version = string(year(today()))
	ls_platform="32"
ELSE
	ls_version = "20"+ string(l_env.pbmajorrevision)+ "." + string(l_env.pbbuildnumber)
	ls_platform=string(l_env.ProcessBitness)
END IF

ls_platform += " Bits"

st_myversion.text=ls_version
st_platform.text=ls_platform

end subroutine

public subroutine wf_probar_python ();//=== wf_probar_python: prueba de humo de PyPb ===
n_cst_pyton lnv_py
string ls_ver, ls_calc

// Un nonvisualobject normal NO se autoinstancia: hay que CREATE-arlo
lnv_py = CREATE n_cst_pyton

// 1) Arrancar el runtime de Python EMBEBIDO y portable (incluye openpyxl)
If lnv_py.of_init("C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\python.runtime\python313.dll") <> 0 Then
	st_pyversion.text = "ERROR al iniciar Python"
	MessageBox("PyPb", "No arranca Python:~r~n" + lnv_py.of_lasterror())
	DESTROY lnv_py
	Return
End If

// 2) Version de Python  ->  platform.python_version()
If lnv_py.of_invoke("platform", "python_version", ls_ver) = 0 Then
	st_pyversion.text = "Python " + ls_ver
Else
	st_pyversion.text = "(no se pudo leer la version)"
	MessageBox("PyPb", "platform.python_version() fallo:~r~n" + lnv_py.of_lasterror())
	DESTROY lnv_py
	Return
End If

// 3) Sanity check ejecutando codigo Python suelto
lnv_py.of_run("2 ** 100", ls_calc)

MessageBox("PyPb OK", "Version: " + ls_ver + "~r~n2 ^ 100 = " + ls_calc)

DESTROY lnv_py
end subroutine

public subroutine wf_probar_excel ();//=== wf_probar_excel: crea un .xlsx con openpyxl sobre el runtime EMBEBIDO ===
n_cst_pyton_excel lnv_xls
string ls_dll, ls_ruta

// Runtime portable que vive DENTRO del proyecto (modelo de despliegue al cliente).
// En produccion: ruta relativa al EXE o configurable, en vez de absoluta.
ls_dll  = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\python.runtime\python313.dll"
ls_ruta = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\salida_pypb.xlsx"

lnv_xls = CREATE n_cst_pyton_excel

If lnv_xls.of_init(ls_dll) <> 0 Then
	MessageBox("Excel/PyPb", "No arranca / no importa openpyxl:~r~n" + lnv_xls.of_lasterror())
	DESTROY lnv_xls
	Return
End If

If lnv_xls.of_nuevo() <> 0 Then
	MessageBox("Excel/PyPb", "of_nuevo:~r~n" + lnv_xls.of_lasterror())
	DESTROY lnv_xls
	Return
End If

If lnv_xls.of_titulo_hoja("PyPb") <> 0 Then
	MessageBox("Excel/PyPb", "of_titulo_hoja:~r~n" + lnv_xls.of_lasterror())
	DESTROY lnv_xls
	Return
End If
lnv_xls.of_set_celda(1, 1, "Hola desde PowerBuilder + openpyxl")
lnv_xls.of_set_celda(2, 1, "Runtime Python embebido y portable")
lnv_xls.of_set_celda(1, 2, "Generado por n_cst_pyton_excel")

If lnv_xls.of_guardar(ls_ruta) <> 0 Then
	MessageBox("Excel/PyPb", "of_guardar:~r~n" + lnv_xls.of_lasterror())
	DESTROY lnv_xls
	Return
End If

MessageBox("Excel/PyPb OK", "Fichero creado:~r~n" + ls_ruta + "~r~n~r~nDiagnostico of_set: " + lnv_xls.of_lastdiag())
DESTROY lnv_xls
end subroutine

public subroutine wf_cargar_facturas ();//=== Carga data2026.json en dw_1 (grid visual) via JSONParser ===
JSONParser ljp
string ls_err, ls_file
long ll_root, ll_item, ll_count, ll_row, i

ls_file = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\data2026.json"

ljp = CREATE JSONParser
ls_err = ljp.LoadFile(ls_file)
If ls_err <> "" Then
	MessageBox("JSON", "No se pudo cargar el JSON:~r~n" + ls_err)
	DESTROY ljp
	Return
End If

ll_root  = ljp.GetRootItem()
ll_count = ljp.GetChildCount(ll_root)

dw_1.SetRedraw(False)
dw_1.Reset()

For i = 1 To ll_count
	ll_item = ljp.GetChildItem(ll_root, i)
	ll_row  = dw_1.InsertRow(0)
	dw_1.SetItem(ll_row, "fecha",       Left(ljp.GetItemString(ll_item, "fecha"), 10))
	dw_1.SetItem(ll_row, "serie",       ljp.GetItemString(ll_item, "serie"))
	dw_1.SetItem(ll_row, "factura",     ljp.GetItemString(ll_item, "factura"))
	dw_1.SetItem(ll_row, "cliente",     ljp.GetItemString(ll_item, "cliente"))
	dw_1.SetItem(ll_row, "razon",       ljp.GetItemString(ll_item, "razon"))
	dw_1.SetItem(ll_row, "cod_fp",      ljp.GetItemString(ll_item, "cod_fp"))
	dw_1.SetItem(ll_row, "forma_pago",  ljp.GetItemString(ll_item, "forma_pago"))
	dw_1.SetItem(ll_row, "subtotal",    ljp.GetItemNumber(ll_item, "subtotal"))
	dw_1.SetItem(ll_row, "total_iva",   ljp.GetItemNumber(ll_item, "total_iva"))
	dw_1.SetItem(ll_row, "importe",     ljp.GetItemNumber(ll_item, "importe"))
	dw_1.SetItem(ll_row, "situacion",   ljp.GetItemString(ll_item, "situacion"))
	dw_1.SetItem(ll_row, "obra",        ljp.GetItemString(ll_item, "obra"))
	dw_1.SetItem(ll_row, "descripcion", ljp.GetItemString(ll_item, "descripcion"))
	dw_1.SetItem(ll_row, "empresa",     ljp.GetItemString(ll_item, "empresa"))
	dw_1.SetItem(ll_row, "anyo",        ljp.GetItemString(ll_item, "anyo"))
Next

dw_1.SetRedraw(True)
DESTROY ljp

st_pyversion.text = String(ll_count) + " facturas cargadas"
end subroutine

public subroutine wf_excel_nativo ();//=== Export NATIVO de PB: SaveDisplayedDataAs(XLSX!) -> .xlsx real, SIN colores/fuentes ===
string ls_path
integer li_rc

If dw_1.RowCount() < 1 Then
	MessageBox("Excel nativo", "Primero carga las facturas.")
	Return
End If

ls_path = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\facturas_nativo.xlsx"

// SaveDisplayedDataAs: .xlsx REAL con los VALORES MOSTRADOS (cabeceras tal como
// se ven, formato de numero/moneda mostrado, computed/totales) y respeta las
// columnas ocultas. PERO no guarda atributos de fuente: SIN colores, negrita ni fills.
li_rc = dw_1.SaveDisplayedDataAs(ls_path, XLSX!)

If li_rc = 1 Then
	MessageBox("Excel nativo (PB XLSX)", "Guardado:~r~n" + ls_path + "~r~n~r~n.xlsx real con datos y cabeceras mostrados, pero SIN colores ni negrita (SaveDisplayedDataAs no guarda fuentes).")
Else
	MessageBox("Excel nativo (PB XLSX)", "Error en SaveDisplayedDataAs: " + String(li_rc))
End If
end subroutine

public subroutine wf_excel_python ();//=== Export PYTHON (openpyxl): cabecera color, zebra, moneda EUR, totales ===
n_cst_pyton lnv_py
n_cst_pypbmodule lnv_mod
n_cst_invocationrequest lnv_req
n_cst_pypbobject lnv_res
string ls_dll, ls_out, ls_in, ls_cols, ls_name
long ll_filas, ll_ncols, j

If dw_1.RowCount() < 1 Then
	MessageBox("Excel Python", "Primero carga las facturas.")
	Return
End If

ls_dll = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\python.runtime\python313.dll"
ls_out = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\facturas_python.xlsx"
ls_in  = "C:\proyecto pw2025\Blog\PowerBuilder\PyPbExample\data2026.json"

// Columnas VISIBLES del DataWindow (en su orden) -> el Excel respeta lo que
// se ve en el grid (ocultar / reordenar columnas).
ll_ncols = Long(dw_1.Describe("DataWindow.Column.Count"))
ls_cols = ""
For j = 1 To ll_ncols
	ls_name = dw_1.Describe("#" + String(j) + ".Name")
	If dw_1.Describe(ls_name + ".Visible") = "1" Then
		If ls_cols <> "" Then ls_cols += ","
		ls_cols += ls_name
	End If
Next

lnv_py = CREATE n_cst_pyton

If lnv_py.of_init(ls_dll) <> 0 Then
	MessageBox("Excel Python", "No arranca Python:~r~n" + lnv_py.of_lasterror())
	DESTROY lnv_py
	Return
End If

If lnv_py.of_import("export_facturas", lnv_mod) <> 0 Then
	MessageBox("Excel Python", "No importa export_facturas:~r~n" + lnv_py.of_lasterror())
	DESTROY lnv_py
	Return
End If

lnv_req = lnv_mod.of_createinvocationrequest("export_file")
lnv_req.of_addargument(ls_in)
lnv_req.of_addargument(ls_out)
lnv_req.of_addargument(ls_cols)

If lnv_mod.of_invoke(lnv_req, lnv_res) <> 0 Then
	MessageBox("Excel Python", "Fallo en export():~r~n" + lnv_mod.of_lasterrormessage())
	DESTROY lnv_py
	Return
End If

ll_filas = lnv_res.of_toint()
DESTROY lnv_py

MessageBox("Excel Python (openpyxl) OK", String(ll_filas) + " filas ->~r~n" + ls_out + &
	"~r~n~r~nCabecera color, filas zebra, moneda EUR, fila de totales, autofiltro y cabecera fija.")
end subroutine

on w_main.create
this.p_2=create p_2
this.st_info=create st_info
this.st_myversion=create st_myversion
this.st_platform=create st_platform
this.r_2=create r_2
this.cb_probar=create cb_probar
this.st_pyversion=create st_pyversion
this.cb_excel=create cb_excel
this.dw_1=create dw_1
this.cb_cargar=create cb_cargar
this.cb_xls_nat=create cb_xls_nat
this.cb_xls_py=create cb_xls_py
this.Control[]={this.p_2,&
this.st_info,&
this.st_myversion,&
this.st_platform,&
this.r_2,&
this.cb_probar,&
this.st_pyversion,&
this.cb_excel,&
this.dw_1,&
this.cb_cargar,&
this.cb_xls_nat,&
this.cb_xls_py}
end on

on w_main.destroy
destroy(this.p_2)
destroy(this.st_info)
destroy(this.st_myversion)
destroy(this.st_platform)
destroy(this.r_2)
destroy(this.cb_probar)
destroy(this.st_pyversion)
destroy(this.cb_excel)
destroy(this.dw_1)
destroy(this.cb_cargar)
destroy(this.cb_xls_nat)
destroy(this.cb_xls_py)
end on

event open;wf_version()
wf_cargar_facturas()






end event

event closequery;
end event

event resize;//=== Reajuste de controles al redimensionar la ventana ===
Integer li_margen = 24

// Banda superior y textos de version/bits (anclados a la derecha)
r_2.width = newwidth
st_myversion.x = newwidth - st_myversion.width - 20
st_platform.x  = newwidth - st_myversion.width - 20

// Pie (copyright) anclado abajo-derecha
st_info.x = newwidth - st_info.width - 20
st_info.y = newheight - st_info.height - 20

// DataWindow de facturas: crece con la ventana. Ancho hasta el borde y
// alto hasta justo encima del pie. Guardas para ventanas pequenas.
If IsValid(dw_1) Then
	If newwidth  > dw_1.x + 450 Then dw_1.width  = newwidth  - dw_1.x - li_margen
	If newheight > dw_1.y + 350 Then dw_1.height = newheight - dw_1.y - 110
End If

end event

type p_2 from picture within w_main
integer x = 5
integer y = 4
integer width = 1253
integer height = 248
boolean originalsize = true
string picturename = "logo.jpg"
boolean focusrectangle = false
end type

type st_info from statictext within w_main
integer x = 5134
integer y = 2532
integer width = 1289
integer height = 52
integer textsize = -7
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 8421504
long backcolor = 553648127
string text = "Copyright © Ramón San Félix Ramón  rsrsystem.soft@gmail.com"
boolean focusrectangle = false
end type

type st_myversion from statictext within w_main
integer x = 3118
integer y = 56
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Versión"
alignment alignment = right!
boolean focusrectangle = false
end type

type st_platform from statictext within w_main
integer x = 3118
integer y = 144
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Bits"
alignment alignment = right!
boolean focusrectangle = false
end type

type r_2 from rectangle within w_main
long linecolor = 33554432
linestyle linestyle = transparent!
integer linethickness = 4
long fillcolor = 33521664
integer width = 3680
integer height = 260
end type

type cb_probar from commandbutton within w_main
integer x = 50
integer y = 644
integer width = 750
integer height = 120
integer taborder = 10
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Probar Python"
end type

event clicked;Parent.wf_probar_python()
end event

type st_pyversion from statictext within w_main
integer x = 50
integer y = 804
integer width = 649
integer height = 84
integer textsize = -12
integer weight = 700
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 134217741
long backcolor = 67108864
boolean focusrectangle = false
end type

type cb_excel from commandbutton within w_main
integer x = 50
integer y = 932
integer width = 750
integer height = 120
integer taborder = 20
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Crear Excel (openpyxl)"
end type

event clicked;Parent.wf_probar_excel()
end event

type dw_1 from datawindow within w_main
integer x = 855
integer y = 296
integer width = 5568
integer height = 2200
integer taborder = 30
string title = "none"
string dataobject = "dw_facturas"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

type cb_cargar from commandbutton within w_main
integer x = 50
integer y = 304
integer width = 750
integer height = 112
integer taborder = 40
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Consultar (JSON)"
end type

event clicked;Parent.wf_cargar_facturas()
end event

type cb_xls_nat from commandbutton within w_main
integer x = 50
integer y = 1276
integer width = 750
integer height = 112
integer taborder = 50
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Excel nativo (PB XLSX)"
end type

event clicked;Parent.wf_excel_nativo()
end event

type cb_xls_py from commandbutton within w_main
integer x = 50
integer y = 1404
integer width = 750
integer height = 112
integer taborder = 60
integer textsize = -10
integer weight = 700
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Excel Python (openpyxl)"
end type

event clicked;Parent.wf_excel_python()
end event

