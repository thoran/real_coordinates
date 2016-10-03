# Module#rmdef

# 2010.04.03
# 0.0.0

# History: Taken from Module#undef 0.3.1.  

class Module
  
  def rmdef(*args)
    args.each do |e|
      self.send(:remove_method, e) if self.instance_methods.include?(e.to_s)
    end
  end
  
end
