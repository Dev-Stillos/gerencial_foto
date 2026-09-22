<%
sub_itens = split(";[[3:1]];[[1:3]];[[20:2]];[[20:2]]",";")
sql = ""

for each item_composto in sub_itens
    if (len(item_composto) > 3) then
        item_composto = replace(item_composto,"[[","")
        item_composto = replace(item_composto,"]]","")
        item_composto = split(item_composto,":")              
        cod_item = item_composto(0)
        qtd = item_composto(1)
        sql = sql & "INSERT INTO tb_est_item_composto VALUES (0,"&cod_item_pai&","&cod_item&","&qtd&")<br>"                   
    
    end if           
next
response.write(sql)


%>