#pragma once
#include <Arduino.h>

struct TelemetryRec {
  uint32_t ts;        // epoch si hay, si no millis()
  uint16_t diMask;    // 12 bits
  int16_t  tempC;
  float    irmsA;
  bool     funciona;
};
