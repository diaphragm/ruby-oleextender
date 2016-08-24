# encoding: utf-8

require 'oleextender'

module VsShape
	OLE_TYPE = "IVShape"
	
end
WIN32OLE::AUTO_EXTEND_MODULES << VsShape



def vscn(prog_id = 'Visio.Application')
	WIN32OLE.connect(prog_id)
end

