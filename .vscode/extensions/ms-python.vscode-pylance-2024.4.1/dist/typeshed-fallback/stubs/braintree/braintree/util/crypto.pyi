text_type = str

class Crypto:
    @staticmethod
    def sha1_hmac_hash(secret_key, content): ...
    @staticmethod
    def sha256_hmac_hash(secret_key, content): ...
    @staticmethod
    def secure_compare(left, right): ...
