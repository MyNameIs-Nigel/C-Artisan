xcount = InputBox("Enter the number of cells on the X")
ycount = InputBox("Enter the number of cells on the Y")
s = msgbox("Default: 0.125 inch border, 0.8 inch cell size", 4,"Custom Dimensions?")

IF s = 6 THEN
cell = InputBox("Enter a cell dimension (assuming it is square)")
cborder = InputBox("Enter the distance between the cells")

ELSE
cell = 0.8
cborder = 0.125
END IF

msgbox("You need to make the inside dimensions " & ( ( xcount*cell ) + ( ( xcount - 1 ) * cborder ) ) & " inches by " & ( ( ycount*cell ) + ( ( ycount - 1 ) * cborder )) & " inches.")