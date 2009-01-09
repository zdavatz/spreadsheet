class Array
  def rcompact
    dup.rcompact!
  end
  def rcompact!
    while !empty? && last.nil?
      pop
    end
    self
  end
end
