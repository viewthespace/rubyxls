$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'rubyxls'

RSpec.configure do |config|

  config.after(:all) do
    FileUtils.rm_rf(tmp_dir_path)
  end

end

def tmp_dir_path
  File.expand_path('../../tmp/', __FILE__)
end
