let V(Volume)

V = 0m3
PartBody.Query("Solid","").Compute("+","Solid","smartVolume(x)",V)

Message("Ìå»ıÊÇ",V)