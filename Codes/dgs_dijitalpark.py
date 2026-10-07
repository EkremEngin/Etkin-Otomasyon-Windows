"""Dijitalpark Teknokent DGS ayarları; giriş ortak doğrulamalı motoru kullanır."""
import dgs_park

PARK = dgs_park.PARKS["Dijitalpark"]


def overrides(_motor):
    """Park için doğrulanmış farklı bir form davranışı bulunursa burada tanımlanır."""


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
