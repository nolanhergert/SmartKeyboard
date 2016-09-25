from Tkinter import *
root = Tk()
#root.image = PhotoImage(file='ripples_sand.gif')
#label = Label(image=root.image, bg='white')
c = Canvas(root,width=500,height=500,bg='blue')
c.pack(expand=YES, fill=BOTH)
m = Label(c, text="This is a test of the splash screen\nThis is only a test.\nwww.sunjay-varma.com",bg='white')
m.pack(side=TOP, expand=YES)
m.config(bg='white', justify=CENTER, font=("calibri", 9))
Button(c, text="Press this button to kill the program", bg='red', command=root.destroy).pack(side=BOTTOM, fill=X)

# Turns off the window border
root.overrideredirect(True)
# Not sure yet
root.geometry("+50+50")
root.lift()
root.wm_attributes("-topmost", True)
#Don't disable input! (for now)
#root.wm_attributes("-disabled", True)
# Windows only, set color to be transparent to white http://www.tcl.tk/man/tcl/TkCmd/wm.htm#M8
root.wm_attributes("-transparentcolor", "white")
m.pack()
c.mainloop()

# Need to add:
# ESC to exit?
# Listen for hotkey, then kill tk window



