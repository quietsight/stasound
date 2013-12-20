<script>
	function change(id, newClass) {
	
	identity=document.getElementById(id);
	
	identity.className=newClass;
	
	}

	function popwin(fileName)
	{
		pcInfoWin = window.open('','InfoWindow','scrollbars=auto,status=no,width=400,height=300')
		pcInfoWin.location.href = fileName;
	}
    
	var tabs = [<%=strTabCnt%>];

	function showTab( tab ){

		// first make sure all the tabs are hidden
		for(i=0; i < tabs.length; i++){
			var obj = document.getElementById(tabs[i]);
			obj.style.display = "none";
		}
			
		// show the tab we're interested in
		var obj = document.getElementById(tab);
		obj.style.display = "block";

	}
</script>
