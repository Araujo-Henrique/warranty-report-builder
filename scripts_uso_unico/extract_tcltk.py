import tkinter

r = tkinter.Tk()
script = r"""
proc copyDir {src dst} {
    file mkdir $dst
    foreach f [glob -nocomplain -directory $src *] {
        set name [file tail $f]
        set target [file join $dst $name]
        if {[file isdirectory $f]} {
            copyDir $f $target
        } else {
            file copy -force $f $target
        }
    }
}
copyDir {//zipfs:/lib/tcl/tcl_library} {C:/Automacao/warranty_app/tcl_data/tcl_library}
copyDir {//zipfs:/lib/tk/tk_library} {C:/Automacao/warranty_app/tcl_data/tk_library}
"""
r.tk.eval(script)
r.destroy()
print("done")
