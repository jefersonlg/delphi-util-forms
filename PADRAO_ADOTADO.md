# delphi-util-forms</br>
Forms that will help in any comercial software</br>
</br>
</br>
</br>
# The contributions may have this format:</br>
</br>
The basic structures and indentation of "begin", "end" and "else", had implicit explained on "if structures".</br>
The indentation is formed by 2 spaces "  ".</br>
</br>
* "if" structures with one condition</br>
</br>
if ( [condition]  ) then</br>
begin</br>
&nbsp;&nbsp;[...]</br>
end</br>
else</br>
&nbsp;&nbsp;begin</br>
&nbsp;&nbsp;&nbsp;&nbsp;[...]</br>
&nbsp;&nbsp;end;</br>
</br>
</br>
with two or more conditions and two or more "if" structures</br>
In this case, the second condition and the others are aligned with first condition</br>
</br>
if ( [condition1]</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[or|and]</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[condition2]</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[or|and]</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[condition3]</br>
) then</br>
begin</br>
&nbsp;&nbsp;[...]</br>
end</br>
else</br>
&nbsp;&nbsp;if ( [condition] ) then</br>
&nbsp;&nbsp;begin</br>
&nbsp;&nbsp;&nbsp;&nbsp;[...]</br>
&nbsp;&nbsp;end;</br>
</br>
</br>
</br>
* "procedure" and "function" structures</br>
</br>
procedure [procedure_name]; | function [function_name] : [type];</br>
</br>
const [const_name] : [type];</br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[more_const] : [type];</br>
</br>
var [var_name] : [type];</br>
&nbsp;&nbsp;&nbsp;&nbsp;[more_var] : [type];</br>
</br>
&nbsp;&nbsp;procedure [inner_procedure_name]; | function [inner_function_name] : [type];</br>
&nbsp;&nbsp;[...]</br>
&nbsp;&nbsp;begin</br>
&nbsp;&nbsp;&nbsp;&nbsp;[...]</br>
&nbsp;&nbsp;end;</br>
</br>
var [local_name_var] : [type];</br>
</br>
begin</br>
&nbsp;&nbsp;[...]</br>
end;</br>
