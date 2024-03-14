from pydantic import BaseModel
from typing import List

class Patients(BaseModel):
    patients: List[int]

class Visits(BaseModel):
    visits: List[int]
    years: List[int]

class RequestedZigzag(BaseModel):
    p_id: int
    visits: List[int]