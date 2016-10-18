<%@ LANGUAGE="JScript"%>
<!-- #include virtual="/_admin/_custom/_catalog/functions.asp" -->
<!-- #include virtual="/_include/functions.asp" -->
<%
try
{
    engineStart();

    var y_price = 0; //dsLoadDocumentPropertyValueByName(catalogDocID, 'yml_price');
    if ( y_price=='' || y_price==null || y_price=='null' || y_price=='undefined' || isNaN(y_price)) y_price = 0;
    
    var groups_ids = new Array();
    var groups = new Array();
    var items = catEnumYandexItems();
    var item, group_id, group;
    var j = 0;
    for (var i in items)
    {
        item = items[i];
        group_id = item.group_id;

        while (group_id && (groups_ids[group_id] != 1))
        {
            group = catGetGroup(group_id);
            groups_ids[group_id] = 1;
            groups[j++] = group;

            group_id = group.parent_id;
        }
    }

	Response.AddHeader( 'Content-Type', 'text/xml; charset=windows-1251' ) ;

    Response.Write('<?xml version="1.0" encoding="windows-1251"?>'+"\r\n");
    Response.Write('<!DOCTYPE yml_catalog SYSTEM "shops.dtd">'+"\r\n");
    
    var date = new Date();
    var day     = date.getDate();
    var month   = date.getMonth()+1;
    var year    = date.getFullYear();
    var hours   = date.getHours();
    var minutes = date.getMinutes();
    var date_str = year+'-'+month+'-'+day+' '+hours+':'+minutes;

    Response.Write('<yml_catalog date="'+date_str+'">'+"\r\n");

    Response.Write('<shop>'+"\r\n");
    var url = getDocURL( catalogDocID );
    var name = filterHTML('tescoma-shop.ru');
    var company = filterHTML('ООО &laquo;Стиль&raquo;');

    Response.Write('<name>'+name+'</name>'+"\r\n");
    Response.Write('<company>'+company+'</company>'+"\r\n");
    Response.Write('<url>'+url+'</url>'+"\r\n");
    Response.Write('<currencies><currency id="RUR" rate="1"/></currencies>'+"\r\n\r\n");

    Response.Write('<categories>'+"\r\n");
    for (var i in groups)
    {
        var group = groups[i];
        var id = filterHTML(group.id);
        var name = filterHTML(group.name);
        var parent_id = filterHTML(group.parent_id);
        if(group.id){
            if (group.parent_id)
                Response.Write('<category id="'+id+'" parentId="'+parent_id+'">'+name+'</category>'+"\r\n");
            else
                Response.Write('<category id="'+id+'">'+name+'</category>'+"\r\n");
        }
    }
    Response.Write('</categories>'+"\r\n\r\n");

    Response.Write('<offers>'+"\r\n");

    var rE = new RegExp ('^.*tescoma.*$','i');
    for (var i in items)
    {
        var item = items[i];
        var id = filterHTML(item.id);
        var group_id = filterHTML(item.group_id);
        if (rE.test(item.name))
            var name = filterHTML(item.name);
        else
            var name = filterHTML('TESCOMA '+item.name);
        var description = filterHTML(item.tech_data);
	var barcode = filterHTML(item.barcode);
        var group_id = filterHTML(item.group_id);
        var available = 'true';
        var url = getGoodLink( item.id );
        var discont = - parseFloat(y_price) / 100;

        var price = Math.round(item.price * (1 - discont) * 100) / 100;
        var price = String(price).replace(/,/g, ".");

        if ( String(item.image_small).length > 0)
            img = Application('HOST_SITE_URL') + Application('CATALOG_IMAGES') + item.image_small;
        else
            img = Application('HOST_SITE_URL') + goodNoSmallImage;
        
        if ( String(item.img_tov).length > 0)
            img_tov = Application('HOST_SITE_URL') + Application('CATALOG_IMAGES')+"imgDop/" + item.img_tov;
        else
            img_tov = '';
            
        if ( String(item.img_col1).length > 0)
            img_col1 = Application('HOST_SITE_URL') + Application('CATALOG_IMAGES')+"imgDop/" + item.img_col1;
        else
            img_col1 = '';
            
        if ( String(item.img_fun1).length > 0)
            img_fun1 = Application('HOST_SITE_URL') + Application('CATALOG_IMAGES')+"imgDop/" + item.img_fun1;
        else
            img_fun1 = '';
            
        if ( String(item.img_det3).length > 0)
            img_det3 = Application('HOST_SITE_URL') + Application('CATALOG_IMAGES')+"imgDop/" + item.img_det3;
        else
            img_det3 = '';
            
        Response.Write('<offer id="'+id+'" available="'+available+'">'+"\r\n");
        Response.Write('<url>'+url+'</url>'+"\r\n");
        Response.Write('<price>'+price+'</price>'+"\r\n");
        Response.Write('<currencyId>RUR</currencyId>'+"\r\n");
        Response.Write('<categoryId>'+group_id+'</categoryId>'+"\r\n");
        Response.Write('<picture>'+img+'</picture>'+"\r\n");
        if(img_tov != '')
        Response.Write('<picture>'+img_tov+'</picture>'+"\r\n");
        if(img_col1 != '')
        Response.Write('<picture>'+img_col1+'</picture>'+"\r\n");
        if(img_fun1 != '')
        Response.Write('<picture>'+img_fun1+'</picture>'+"\r\n");
        if(img_det3 != '')
        Response.Write('<picture>'+img_det3+'</picture>'+"\r\n");
        Response.Write('<name>'+name+'</name>'+"\r\n");
        Response.Write('<description>'+description+'</description>'+"\r\n");
        if (barcode != '') Response.Write('<barcode>'+barcode+'</barcode>'+"\r\n");
        Response.Write('</offer>'+"\r\n\r\n");
    }
    Response.Write('</offers>'+"\r\n");

    Response.Write('</shop>'+"\r\n");

    Response.Write('</yml_catalog>');

//    Response.Write('Generation of YML file is done!');
}
catch(e)
{
        Response.Write(e.description);
        //Server.Transfer("/_errors/er500.asp");
}
finally
{
        engineStop();
}

%>