// NicoModBus.h - Biblioteca simplificada para Modbus TCP
//#include <ESP8266WiFi.h>
//#include <WiFiClient.h>

class NicoModBus {
private:
    WiFiClient client;
    IPAddress serverIP;
    uint16_t serverPort;
    uint8_t unitID;
    
public:
    NicoModBus() : serverPort(502), unitID(1) {}
    
    bool connect(IPAddress ip, uint16_t port = 502) {
        serverIP = ip;
        serverPort = port;
        return client.connect(ip, port);
    }
    
    void disconnect() {
        client.stop();
    }
    
    bool readDiscreteInputs(uint16_t offset, bool* values, uint8_t count) {
        if (!client.connected()) {
            if (!connect(serverIP, serverPort)) return false;
        }
        
        // Construir trama Modbus TCP
        uint8_t request[12] = {
            0x00, 0x01, // Transaction ID (puede ser aleatorio)
            0x00, 0x00, // Protocol ID (0 para Modbus)
            0x00, 0x06, // Length
            unitID,     // Unit ID
            0x02,       // Function Code (Read Discrete Inputs)
            highByte(offset), lowByte(offset), // Starting Address
            highByte(count), lowByte(count)    // Quantity
        };
        
        client.write(request, sizeof(request));
        client.flush();
        
        // Esperar respuesta
        unsigned long start = millis();
        while (client.available() < 9 && millis() - start < 2000) {
            delay(10);
        }
        
        if (client.available() < 9) return false;
        
        // Leer encabezado
        for (int i = 0; i < 7; i++) client.read();
        
        uint8_t functionCode = client.read();
        if (functionCode != 0x02) return false;
        
        uint8_t byteCount = client.read();
        if (byteCount != (count + 7) / 8) return false;
        
        // Leer valores
        uint8_t byteVal;
        for (uint8_t i = 0; i < count; i++) {
            if (i % 8 == 0) byteVal = client.read();
            values[i] = (byteVal >> (i % 8)) & 0x01;
        }
        
        return true;
    }
};
