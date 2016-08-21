unit lltmmudxp;

interface

function  CalcExpNeeded(const Level, Chart: integer): Int64; stdcall;

implementation

function  CalcExpNeeded(const Level, Chart: integer): Int64; stdcall; external 'lltmmudxp.dll';

end.
