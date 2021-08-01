import { ComponentFixture, TestBed } from '@angular/core/testing';

import { StageSettingsComponent } from './stage-settings.component';

describe('StageSettingsComponent', () => {
  let component: StageSettingsComponent;
  let fixture: ComponentFixture<StageSettingsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ StageSettingsComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(StageSettingsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
