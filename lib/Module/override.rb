# Module#override

# 2010.04.04
# 0.0.0

# Description: Works just like Module#include except that it pre-emptively removes any overlapping methods in the target, so as the included methods override the target's methods.  

require 'Module/rmdef'

class Module
  
  def override(*modules)
    modules.reverse_each do |modul|
      modul.instance_methods.each{|m| self.rmdef(m)}
      self.send(:include, modul)
    end
  end
  private :override
  alias_method :include_with_override, :override
  
end
