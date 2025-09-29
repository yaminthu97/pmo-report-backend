class JsonWebToken
  SECRET_KEY = Rails.application.secrets.secret_key_base || ENV['SECRET_KEY_BASE'] 

  def self.encode(payload, exp = 1.hours.from_now)
    payload[:exp] = exp.to_i
    JWT.encode(payload, SECRET_KEY) 
  end

  def self.decode(token)
    body = JWT.decode(token, SECRET_KEY)[0]
    HashWithIndifferentAccess.new body
  rescue
    nil
  end
end