forward
global type n_cst_pyton_excel from nonvisualobject
end type
end forward

global type n_cst_pyton_excel from nonvisualobject
end type
global n_cst_pyton_excel n_cst_pyton_excel

type variables
//===========================================================================
// n_cst_pyton_excel  -  Generacion de .xlsx con openpyxl via PyPb
// USA n_cst_pyton por dentro y expone metodos de negocio (estilo n_cst_venped_pdf).
// Sin excepciones hacia fuera: codigo 0/-1 + of_lasterror().
//===========================================================================
n_cst_pyton        inv_py
n_cst_pypbmodule   inv_openpyxl
n_cst_pypbobject   inv_wb          // Workbook
n_cst_pypbobject   inv_ws          // hoja activa
string             is_lasterror
string             is_diag         // diagnostico del intento of_set (aprendizaje)
end variables

forward prototypes
public function integer of_init (string as_pythonpath)
public function integer of_nuevo ()
public function integer of_titulo_hoja (string as_titulo)
public function integer of_set_celda (long al_fila, long al_col, string as_valor)
public function integer of_guardar (string as_ruta)
public function string of_lasterror ()
public function string of_lastdiag ()
end prototypes

public function integer of_init (string as_pythonpath);/// of_init
///
/// Arranca el runtime de Python (delegando en n_cst_pyton) e importa openpyxl.
///
/// string as_pythonpath: ruta al pythonXXX.dll ("" = runtime local).
///
/// returns: 0 si OK, -1 si error (ver of_lasterror()).

is_lasterror = ""

inv_py = CREATE n_cst_pyton

If inv_py.of_init(as_pythonpath) <> 0 Then
	is_lasterror = inv_py.of_lasterror()
	Return -1
End If

If inv_py.of_import("openpyxl", Ref inv_openpyxl) <> 0 Then
	is_lasterror = inv_py.of_lasterror()
	Return -1
End If

Return 0
end function

public function integer of_nuevo ();/// of_nuevo
///
/// Crea un Workbook nuevo y se queda con la hoja activa.
///
/// returns: 0 si OK, -1 si error.

is_lasterror = ""

// openpyxl.Workbook()  ->  instanciar la clase del modulo
If inv_openpyxl.of_instantiate("Workbook", Ref inv_wb) <> 0 Then
	is_lasterror = inv_openpyxl.of_lasterrormessage()
	Return -1
End If

// wb.active  ->  hoja activa (atributo/propiedad)
If inv_wb.of_getmember("active", Ref inv_ws) <> 0 Then
	is_lasterror = inv_wb.of_lasterrormessage()
	Return -1
End If

Return 0
end function

public function integer of_titulo_hoja (string as_titulo);/// of_titulo_hoja
///
/// Pone el titulo de la hoja activa (ws.title = as_titulo).
///
/// NOTA: of_executestatement usa eval() (solo EXPRESIONES). Una asignacion como
/// "ws.title = t" da SyntaxError (invalid syntax). Por eso usamos setattr(ws,'title',t),
/// que es una EXPRESION (llamada a funcion) y SI hace la asignacion. Guardamos en
/// is_diag lo que devolvio of_set, para aprendizaje.
///
/// returns: 0 si OK, -1 si error.

n_cst_invocationrequest lnv_req
n_cst_pypbobject lnv_res
integer li_set

is_lasterror = ""

// --- Diagnostico: intento por property con of_set ---
li_set = inv_ws.of_set("title", as_titulo)
is_diag = "of_set('title') -> " + String(li_set)
If li_set <> 0 Then is_diag = is_diag + " | err=" + inv_ws.of_lasterrormessage()

// --- Via robusta: setattr(ws,'title',t) (eval acepta llamadas, no asignaciones) ---
lnv_req = inv_ws.of_createinvocationrequest("stub")
lnv_req.of_addnamedargument("ws", inv_ws)
lnv_req.of_addnamedargument("t", as_titulo)

If inv_py.of_exec_req("setattr(ws, 'title', t)", lnv_req, lnv_res) <> 0 Then
	is_lasterror = inv_py.of_lasterror()
	Return -1
End If

Return 0
end function

public function integer of_set_celda (long al_fila, long al_col, string as_valor);/// of_set_celda
///
/// Escribe un valor en una celda  ->  ws.cell(row=, column=, value=)
/// Demuestra el uso de n_cst_invocationrequest con argumentos NOMBRADOS (kwargs).
///
/// long al_fila: fila (1..n)
/// long al_col: columna (1..n)
/// string as_valor: valor a escribir
///
/// returns: 0 si OK, -1 si error.

n_cst_invocationrequest lnv_req
n_cst_pypbobject lnv_cell

is_lasterror = ""

lnv_req = inv_ws.of_createinvocationrequest("cell")
lnv_req.of_addnamedargument("row", al_fila)
lnv_req.of_addnamedargument("column", al_col)
lnv_req.of_addnamedargument("value", as_valor)

If inv_ws.of_invoke(lnv_req, Ref lnv_cell) <> 0 Then
	is_lasterror = inv_ws.of_lasterrormessage()
	Return -1
End If

Return 0
end function

public function integer of_guardar (string as_ruta);/// of_guardar
///
/// Guarda el libro en disco  ->  wb.save(as_ruta)
/// Demuestra n_cst_invocationrequest con argumento POSICIONAL.
///
/// string as_ruta: ruta completa del .xlsx
///
/// returns: 0 si OK, -1 si error.

n_cst_invocationrequest lnv_req
n_cst_pypbobject lnv_res

is_lasterror = ""

lnv_req = inv_wb.of_createinvocationrequest("save")
lnv_req.of_addargument(as_ruta)

If inv_wb.of_invoke(lnv_req, Ref lnv_res) <> 0 Then
	is_lasterror = inv_wb.of_lasterrormessage()
	Return -1
End If

Return 0
end function

public function string of_lasterror ();Return is_lasterror
end function

public function string of_lastdiag ();Return is_diag
end function

on n_cst_pyton_excel.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_pyton_excel.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

event destructor;// El contexto Python es global (PyPb lo reutiliza); destruir el helper
// NO termina Python, solo libera este objeto.
If IsValid(inv_py) Then DESTROY inv_py
end event
