<template>
  <div class="dynamic-table">
    <div class="head-com">
      <div class="title"><span>任务清单</span>
        <span class="el-icon-close close-btn"></span>
      </div>
      <div class="diviv"></div>
      <div class="switch-com">
        <div class="btn"
             :class="{active:tableActive==='base'}"
             @click="switchTable('base')"><span>小程序清单</span></div>
        <div class="btn"
             :class="{active:tableActive==='fun'}"
             @click="switchTable('fun')"><span>app清单</span></div>
      </div>
    </div>
    <div class="context">
      <excelToTableCom ref="exceldom" />
    </div>
  </div>
</template>

<script>
import excelToTableCom from './excelToTableCom.vue'
export default {
  data () {
    return {
      showpop: false,
      tableActive: 'base'
    }
  },
  components: { excelToTableCom },
  methods: {
    switchTable (type) {
      this.tableActive = type;
      this.$refs.exceldom.switchTableData(type)
    },
    handleClose (done) {
      this.$confirm('确认关闭？')
        .then(_ => {
          done();
        })
        .catch(_ => { });
    }
  }
}
</script>
<style lang="scss" scoped>
.dynamic-table {
  position: relative;
  left: 0px;
  right: 0px;
  margin: auto;
  z-index: 9999;
  width: 100%;
  height: 100%;
  overflow: hidden;
  overflow-x: auto;
  border: 1px solid #99edff;
  background: rgba(1, 20, 37, 0.9);
  .head-com {
    position: relative;
    .close-btn {
      position: absolute;
      top: 15px;
      right: 20px;
      font-size: 30px;
      color: #fff;
      cursor: pointer;
    }
    .title {
      line-height: 65px;
      font-size: 26px;
      font-weight: 700;
      text-align: center;
      color: #fff;
    }
    .diviv {
      width: 98%;
      height: 1px;
      margin: 10px auto;
      border: 1px solid #476280;
      position: relative;
      &::after {
        content: "";
        position: absolute;
        left: 0;
        top: -1px;
        width: 36px;
        height: 2px;
        background: rgba(184, 220, 239, 1);
      }
      &::before {
        content: "";
        position: absolute;
        right: 0;
        top: -1px;
        width: 36px;
        height: 2px;
        background: rgba(184, 220, 239, 1);
      }
    }
    .switch-com {
      display: flex;
      justify-content: center;
      height: 50px;
      align-content: center;
      font-size: 24px;
      color: #fff;
      .btn {
        position: relative;
        width: 200px;
        height: 30px;
        cursor: pointer;
        text-align: center;
        background: linear-gradient(
          90deg,
          #173e63 0%,
          rgba(16, 63, 103, 0.5) 100%
        );

        &.active {
          background: linear-gradient(
            90deg,
            #2984d9 0%,
            rgba(65, 135, 194, 0.5) 100%
          );
        }
        &.active::after {
          content: "";
          position: absolute;
          top: 10px;
          left: 30px;
          width: 10px;
          height: 10px;
          border-radius: 50%;
          background: rgba(253, 255, 132, 1);
        }
        span {
          font-size: 16px;
          font-weight: 400px;
          color: #fff;
        }
      }
    }
  }
  .context {
    margin: 0 auto;
    width: 98%;
    height: 950px;
    color: #fff;
    font-size: 24px;
    overflow: hidden;
    overflow-y: auto;
    overflow-x: auto;

    /* 滚动条整体样式 */
    &::-webkit-scrollbar {
      width: 8px; /* 竖直滚动条宽度 */
      height: 8px; /* 水平滚动条高度 */
    }

    /* 滚动条滑块 */
    &::-webkit-scrollbar-thumb {
      border-radius: 8px; /* 圆角滑块 */
      background: rgba(34, 99, 160, 0.5); /* 设置滑块颜色 */
    }
    /* 滚动条轨道（背景） */
    &::-webkit-scrollbar-track {
      background: transparent; /* 设置轨道颜色 */
    }
    &::-webkit-scrollbar-track-piece {
      background-color: transparent; /* 设置轨道颜色 */
    }
  }
}
</style>