/***********************************************
Esta função equivale a função Trim() do VBScript
***********************************************/
function jTrim(strMsg)
{   
	strNm = '';
	tamstr = strMsg.length;
	tot = tamstr;
	ileft = 0;
	iright= 0;
	
	while(strMsg.charAt(ileft)== " ")
	{
		ileft = ileft + 1;
	}

	if (ileft == tot)
	{
		return "";
	}
				
	while(strMsg.charAt(tamstr-1)== " ")
	{
		tamstr = tamstr - 1;
		iright = iright + 1;
	}

	strNm = strMsg.substring(ileft,tot-iright);
	return strNm;	
}

function CaracteresValidos(CharsValidos,str) {
// Retorna TRUE se o string str for composto somente
// por caracteres contidos no string CharsValidos.
	str = str.toUpperCase();
	var fc_i;
	var fc_char1;
	var flag;
	var fc_a=str.length - 1;

	if(jTrim(str)=="")
	{
		return;
	}

	for (fc_i=0; fc_i<=fc_a; fc_i++) {
		fc_char1 = str.substring(fc_i, fc_i+1)
		flag=CharsValidos.indexOf(fc_char1)
		if (flag==-1){
			document.activeElement.setAttribute('value', str.substr(0, str.indexOf(fc_char1)));
			return;
		}
	}
	return(true);
}
