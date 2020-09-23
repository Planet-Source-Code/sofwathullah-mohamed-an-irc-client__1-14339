Attribute VB_Name = "errorMesgs"

'to make access to these codes easier
'now you only have to type err. and VB will display the list
'for ya. wowee :) - simon
'ps. we'll add more as we go along

Public Enum err
    ERR_NICKNAMEINUSE = 433
    ERR_NICKCOLLISION = 436
    RPL_TOPIC = 332
    RPL_NAMREPLY = 353
    RPL_ENDOFNAMES = 366
End Enum


