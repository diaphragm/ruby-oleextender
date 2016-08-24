# encoding: utf-8

require 'oleextender'

module ACADApplication
  OLE_TYPE = "IAcadApplication"
  
  def ad; ActiveDocument() end
end

module ACADDocument
  OLE_TYPE = "IAcadDocument"
  
  def ass; ActiveSelectionSet() end
end


def accn(prog_id = 'AutoCAD.Application')
	WIN32OLE.connect(prog_id)
end

