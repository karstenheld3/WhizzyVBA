@ECHO OFF
git add -u
git add .
rem ECHO Enter Commit Comment:
SET /P "COMMENT=Enter Comment: "
IF 	"%COMMENT%"=="" (
	SET COMMENT="New pushed version"
)
git commit -m "%COMMENT%"
git push https://github.com/karstenheld3/WhizzyVBA.git master
pause