// RS256 JWT verification in plain JavaScript (BigInt), for the PocketBase JSVM.

const SHA256_DIGEST_INFO_PREFIX = "3031300d060960864801650304020105000420"
const BASE64URL_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"

function base64UrlToBytes(input) {
    const clean = input.replace(/=+$/, "")
    const bytes = []
    let buffer = 0
    let bitCount = 0
    for (const character of clean) {
        const value = BASE64URL_ALPHABET.indexOf(character)
        if (value < 0) {
            throw new Error("invalid base64url character")
        }
        buffer = (buffer << 6) | value
        bitCount += 6
        if (bitCount >= 8) {
            bitCount -= 8
            bytes.push((buffer >> bitCount) & 0xff)
        }
    }
    return bytes
}

function base64UrlToString(input) {
    const bytes = base64UrlToBytes(input)
    const percentEncoded = bytes.map((byte) => "%" + byte.toString(16).padStart(2, "0")).join("")
    return decodeURIComponent(percentEncoded)
}

function bytesToHex(bytes) {
    return bytes.map((byte) => byte.toString(16).padStart(2, "0")).join("")
}

function bytesToBigInt(bytes) {
    return BigInt("0x" + (bytesToHex(bytes) || "0"))
}

function modularPower(base, exponent, modulus) {
    let result = 1n
    base %= modulus
    while (exponent > 0n) {
        if (exponent & 1n) {
            result = (result * base) % modulus
        }
        exponent >>= 1n
        base = (base * base) % modulus
    }
    return result
}

function expectedPaddedHex(signingInput, modulusByteLength) {
    const digestHex = $security.sha256(signingInput)
    const digestInfoHex = SHA256_DIGEST_INFO_PREFIX + digestHex
    const paddingLength = modulusByteLength - digestInfoHex.length / 2 - 3
    if (paddingLength < 8) {
        throw new Error("modulus too short")
    }
    return "0001" + "ff".repeat(paddingLength) + "00" + digestInfoHex
}

function decode(token) {
    const parts = String(token).split(".")
    if (parts.length !== 3) {
        throw new Error("malformed token")
    }
    return {
        header: JSON.parse(base64UrlToString(parts[0])),
        payload: JSON.parse(base64UrlToString(parts[1])),
        signingInput: parts[0] + "." + parts[1],
        signature: base64UrlToBytes(parts[2]),
    }
}

function verifySignature(decoded, jsonWebKey) {
    if (decoded.header.alg !== "RS256") {
        throw new Error("unexpected alg " + decoded.header.alg)
    }
    const modulusBytes = base64UrlToBytes(jsonWebKey.n)
    const modulus = bytesToBigInt(modulusBytes)
    const exponent = bytesToBigInt(base64UrlToBytes(jsonWebKey.e))
    const signature = bytesToBigInt(decoded.signature)
    if (signature >= modulus) {
        throw new Error("signature out of range")
    }
    const modulusByteLength = modulusBytes[0] === 0 ? modulusBytes.length - 1 : modulusBytes.length
    const actualHex = modularPower(signature, exponent, modulus).toString(16).padStart(modulusByteLength * 2, "0")
    if (actualHex !== expectedPaddedHex(decoded.signingInput, modulusByteLength)) {
        throw new Error("invalid signature")
    }
}

function verifyClaims(payload, expected) {
    const now = Math.floor(Date.now() / 1000)
    const clockSkewSeconds = 300
    if (typeof payload.exp !== "number" || payload.exp + clockSkewSeconds < now) {
        throw new Error("token expired")
    }
    if (typeof payload.nbf === "number" && payload.nbf - clockSkewSeconds > now) {
        throw new Error("token not yet valid")
    }
    if (!expected.audiences.includes(payload.aud)) {
        throw new Error("unexpected audience " + payload.aud)
    }
    if (!expected.issuers.includes(payload.iss)) {
        throw new Error("unexpected issuer " + payload.iss)
    }
    if (payload.tid !== expected.tenantId) {
        throw new Error("unexpected tenant " + payload.tid)
    }
}

// keys: array of JWKs; expected: {audiences, issuers, tenantId}
function verifyRs256(token, keys, expected) {
    const decoded = decode(token)
    const jsonWebKey = keys.find((key) => key.kid === decoded.header.kid)
    if (!jsonWebKey) {
        throw new Error("unknown signing key " + decoded.header.kid)
    }
    verifySignature(decoded, jsonWebKey)
    if (expected) {
        verifyClaims(decoded.payload, expected)
    }
    return decoded.payload
}

module.exports = { decode, verifyRs256, base64UrlToBytes }
