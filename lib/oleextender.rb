# encoding: utf-8

=begin
==何これ
win32oleインスタンスに、moduleを自動でextendする機能を提供するライブラリです。

==前置き
win32oleでExcelなどを操作している場合に、(Excelの)Rangeオブジェクトや
(Excelの)Worksheetオブジェクトに対してメソッドを定義したい場合はありませんか？

Rubyのオブジェクトの場合はclassにインスタンスメソッドを定義してやればOKですが、
(Excelの)Rangeオブジェクトも(Excelの)Worksheetオブジェクトも、Ruby上では全て
WIN32OLEクラスなので、クラスにインスタンスメソッドを定義する方法は使えません。

その場合、メソッドを定義したmoduleを用意して、WIN32OLEクラスのインスタンスに対して、
extendする方法がありますが、いちいち生成したインスタンス全てに対してextendするのは
非常に手間です。

このライブラリは、win32oleインスタンスに、moduleを自動でextendする機能を提供します。

==使い方
  Module Foo
    OLE_TYPE = /Bar/i
  end
  WIN32OLE::AUTO_EXTEND_MODULES << Foo
  
とすると、OLE_TYPE === win32ole.ole_type.name がtrueとなるwin32oleインスタンスに対して、
method_missingが呼び出されたタイミングで自動的にFooがextendされます。
=end



require 'win32ole'

class WIN32OLE
  AUTO_EXTEND_MODULES = []
  
  # win32oleインスタンスに適したモジュールを返す。
  def self.select_module(oleobj)
    AUTO_EXTEND_MODULES.select{|mod|
      f = true
      f &= mod::OLE_TYPE === oleobj.ole_type.name if mod.constants.include?(:OLE_TYPE)
      f &= mod::OLE_TYPELIB === oleobj.ole_typelib.name if mod.constants.include?(:OLE_TYPELIB)
      f &= mod::GUID === oleobj.ole_type.guid if mod.constants.include?(:GUID)
      f
    }
  end
  
  alias :__method_missing :method_missing
  
  def method_missing(name, *args, &block)
    if @flag_auto_extended
      __method_missing(name, *args, &block)
    else
      auto_extend
      __send__(name, *args, &block)
    end
  end
  
  def respond_to_missing?(symbol, include_private)
    if @flag_auto_extended
      super
    else
      auto_extend
      respond_to?(symbol, include_private)
    end
  end
  
  def auto_extend
    unless (wrap_module = self.class.select_module(self)).empty?
      extend(*wrap_module)
    end
    @flag_auto_extended = true
    self
  end
end



# OLE instance Wrapper Moduleの定義

module Global
  alias :with :instance_exec
end
WIN32OLE::AUTO_EXTEND_MODULES << Global


# example

if __FILE__ == $0
  # 匿名モジュールでもOK
  WIN32OLE::AUTO_EXTEND_MODULES << Module.new do
    self::OLE_TYPE = "Dummy"
    def foobar; end
  end
  
end
