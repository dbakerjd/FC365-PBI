import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CreateForecastCycleComponent } from './create-forecast-cycle.component';

describe('CreateForecastCycleComponent', () => {
  let component: CreateForecastCycleComponent;
  let fixture: ComponentFixture<CreateForecastCycleComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ CreateForecastCycleComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(CreateForecastCycleComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
